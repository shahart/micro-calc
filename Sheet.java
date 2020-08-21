/*
 * MicroCalc -- j2me spreadsheet
 *
 * Copyright (c) 2002-2003 Michael Zemljanukha (support@wapindustrial.com)
 *
 * This program is free software; you can redistribute it and/or modify it
 * under the terms of the GNU General Public License as published by the Free
 * Software Foundation; either version 2 of the License, or (at your option)
 * any later version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
 * more details.
 *
 * You should have received a copy of the GNU General Public License along
 * with this program; if not, write to the Free Software Foundation, Inc.,
 * 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 */
/*
 * All the sheet operations without drawing ones
 */

/*
        to do: recalculation strategy should be changed to use wave algorithm
        (level depth) to avoid unnecessary recalculations
                the same with recalculation when inserting/deleting rows/columns
                the crossreference table should have been used
 */

package com.wapindustrial.calc;

//import javax.microedition.rms.*;
import java.util.*;
import java.io.*;

public final class Sheet {
    
    public static final int VERSION		= 0x000708;         // this is not exact 0.7.7 version since I have "reengeenered" it from 0.8.0 that used Float class
    public static final int MAX_FORMULA		= 256;
    public static final int MAX_REFERENCES	= 20;	// in a cell formula
    
    private static final int HEADER_SIZE	= (5*2+4);
    
    public static final String DEFAULT_NAME	= "untitled";
    public static final String SHEET_PREFIX	= "ff_";		// sheet prefix for name  to keep in RecordStore
    
    public int rows, columns;                       // sheet size
    public int serverVersion;                       // version info, higher part is server version, lower is local version
                                                    // when sheet comes to server its serverversion increments and local version (lower part) = 0
                                                    // each time sheet is saves to RS its local version is incremented
    
    public static int newRows=32, newColumns=8;	// to create new sheets, loaded from profile

    public  Hashtable cells = new Hashtable();
    
    private short links[][];	// to keep dependies crossreference table (range to range)
    private int nLinks;
    
    public boolean changed;
    
    public String name;			// sheet name
    
    public short columnWidth[], rowHeight[];
    public int defaultWidth, defaultHeight;
    
//    public Hashtable names;

    // dummy constructor, to use before loading (ugly, fix it)
    public Sheet() {
        this(0,0,0,0);
    }
    
    public Sheet( int _rows, int _columns, int _defaultWidth, int _defaultHeight ) {
        
        rows = _rows; columns = _columns;
        defaultWidth = _defaultWidth; defaultHeight = _defaultHeight;
        
        allocate();
        
        clearSheet();
        
        name = DEFAULT_NAME;
        
    }
    
    public void allocate() {
        columnWidth = new short[columns];
        rowHeight = new short[rows];
    }
    
    public void clearSheet() {
        cells.clear();
        links =  null;
        nLinks = 0;
        changed = false;
        // set default formats
        for( int i=0; i<rows; i++ )
            rowHeight[i] = (short) defaultHeight;
        for( int j=0; j<columns; j++ )
            columnWidth[j] = (short) defaultWidth;
    }
    
    /* ============================================================
        Save/Restore
       ============================================================
     */
    public void loadSheet( String _name, InputStream is )  throws IOException, BadFormulaException {
        DataInputStream dinp = null;
        int nn;
        
        clearSheet();
        
        name = _name;
        
        try {
//            binp = new ByteArrayInputStream( dd ); 
            dinp = new DataInputStream( is );
            // don't read old versions
            int ver = dinp.readShort();
            if( ver <= 0x000703 ) {
                throw new IOException("too old sheet version, cannot open the sheet");
            }
            
            serverVersion = dinp.readInt();
            rows = dinp.readInt();
            columns = dinp.readInt();
            if( rows <= 0 || columns <=0 )
                throw new IOException();
            
            // reallocate memory
            allocate();
            
            int nrec = dinp.readInt();
            
            // columns & rows width
            // to do: height and width must be recalculated in new/old font proportions
            //
            for( int i=0; i<rows; i++ )
                rowHeight[i] = dinp.readShort();
            for( int j=0; j<columns; j++ )
                columnWidth[j] = dinp.readShort();
            
            // 1st pass, load all values
            for( int i=0; i<nrec; i++ ) {
                Result rr = Result.restoreCell( dinp, ver );       // cell
                cells.put( rr, rr );                                // into hashtable
                addReferences( Result.Evaluate( rr.str ), rr.i1, rr.j1 );
            }

        }
        finally {
            if( dinp != null ) dinp.close();
        }
        
    }

    public void saveSheet( OutputStream os ) throws IOException {
        DataOutputStream dout = null;
        
        try {
            // write workbook header
            dout = new DataOutputStream( os );
            dout.writeShort( VERSION );
            
//            serverVersion++;
            dout.writeInt( serverVersion );
            
            // write a sheet
            dout.writeInt( rows );
            dout.writeInt( columns );
            dout.writeInt( cells.size() );
            // rows height
            for( int i=0; i<rows; i++ )
                dout.writeShort( rowHeight[i] );
            // columns height
            for( int j=0; j<columns; j++ )
                dout.writeShort( columnWidth[j] );
            // the sheet cells
            for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
                Result rr = (Result) cll.nextElement();
                rr.save( dout, false );
            }
        }
        finally {
            if( dout != null ) dout.close();
        }
    }
    
    /* =========================================
     * Misc
     */
    public boolean isEmpty( int i, int j ) {
        return getCell( i,j ).isEmptyCell();      // don't compare with CELL_EMPTY since there may be an expression
    }
    public boolean isFormula( int i, int j ) {
        return getCell( i,j ).isFormula();
    }
// must be replace by !isEmpty()
//    public boolean hasFormula( int i, int j ) {
//        return formulas[i][j] != null;
//    }
//    
    
    public int cellY( int I ) {
        int Y = 0;
        for( int i=0; i<I; i++ )
            Y += rowHeight[i];
        return Y;
    }
    
    public int cellX( int J ) {
        int X = 0;
        for( int i=0; i<J; i++ )
            X += columnWidth[i];
        return X;
    }
    
    /* =================================================================
     * references
     */
    // point p2 depends on range r1
    private void addReference( int r1_i1, int r1_j1, int r1_i2, int r1_j2, int p2_i, int p2_j ) {
        
        // test if it's already in the list
        for( int i=0; i<nLinks; i++ )
            if( links[i][0] == r1_i1 && links[i][1] == r1_j1 &&
            links[i][2] == r1_i2 && links[i][3] == r1_j2 &&
            links[i][4] == p2_i && links[i][5] == p2_j )
                return;
        
        // extend array
        if( links == null || links.length < nLinks+1 ) {
            short _links[][] = new short [nLinks+1][6];
            for( int i=0; i<nLinks; i++ ) {
                _links[i][0] = links[i][0];
                _links[i][1] = links[i][1];
                _links[i][2] = links[i][2];
                _links[i][3] = links[i][3];
                _links[i][4] = links[i][4];
                _links[i][5] = links[i][5];
            }
            links = _links;
        }
        
        links[nLinks][0] = (short) r1_i1;
        links[nLinks][1] = (short) r1_j1;
        links[nLinks][2] = (short) r1_i2;
        links[nLinks][3] = (short) r1_j2;
        links[nLinks][4] = (short) p2_i;
        links[nLinks][5] = (short) p2_j;
        nLinks++;
    }
    
    // to do - replace links with hashtable with Result range with p2 in ll, or p2 range in funcargs[0]
    // to do - replace I,J by Cell object
    public void clear( int I, int J ) {
        // remove references from this cell
        for( int i=0; i<nLinks; i++ )
            if( links[i][4] == I && links[i][5] == J ) {
                for( int j=i; j<nLinks-1; j++ ) {
                    links[j][0] = links[j+1][0];
                    links[j][1] = links[j+1][1];
                    links[j][2] = links[j+1][2];
                    links[j][3] = links[j+1][3];
                    links[j][4] = links[j+1][4];
                    links[j][5] = links[j+1][5];
                }
                nLinks--;
            }
        cells.remove( new Integer( (I<<16)|J ) );
    }
    
    // returns true if value has been changed
    // does not recalculate dependencies!
    public void setFormula( int I, int J, String ss ) throws BadFormulaException {
        
        // evaluate the expression
        Result rr = Result.Evaluate( ss );
        Result rrr = rr.calculate( this );
        // no exception occured
        
        clear( I, J );
        
        Result cell = Result.createCell( I,J, ss, rrr );
        cells.put( cell, cell );
        
        addReferences( rr, I, J );
        
    }

    // does not recalculate dependencies!
    public void setFormula1( int i, int j, String ss ) {
        try {
            setFormula( i, j, ss );
            changed = true;
        }
        catch( Exception e ) { }
//        catch( BadFormulaException e ) {
//            System.out.println("system error - exception <" + e.getMessage() + "> wasn't expected here");
//        }
    }
    
    private void addReferences( Result rr, int I, int J ) {
        if( rr.type == Result.TYPE_RANGE ) {
            // p2 depends on p1!
//System.out.println("dep "+rr.i1+" "+rr.j1+" "+rr.i2+" "+rr.j2+" "+I+" "+J);            
            addReference( rr.i1,rr.j1, rr.i2,rr.j2, I,J );
            return;
        }
        for( int i=0; i<rr.funcargs.length; i++ )
            addReferences( rr.funcargs[i], I, J );
    }
    
    // recursive
    public void calculateDepended( int I, int J, boolean first ) {
        
        if( first ) {
            // clear calculated flag (to check loops)
            for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
                Result rr = (Result) cll.nextElement();
                rr.absolute &= ~Result.CALCULATED;
            }
        }
        
        Result cll = getCell( I,J );
        
        // a loop test
        if( (cll.absolute & Result.CALCULATED) != 0 ) {
            cll.funcargs[0] = Result.RESULT_ERROR;
            return;
        }
        
        // for the root it's already calculated and is called when the value was changed
        if( !first ) {
            Result rr = Result.RESULT_ERROR;
            try {
                rr = Result.Evaluate( cll.str ).calculate( this );
            }
            catch( Exception e ) {
            }
            if( cll.funcargs[0].equals( rr ) ) return;	// no change
            
            // changed, recursive calculate depended cells
            cll.funcargs[0] = rr;
        }
        
        cll.absolute |= Result.CALCULATED;
        
        // look trough crossref table to see depended cells
        for( int i=0; i<nLinks; i++ ) {
            if( I >= links[i][0] && I <= links[i][2] &&
            J >= links[i][1] && J <= links[i][3] ) {
                //System.out.println("calculating");
                calculateDepended( links[i][4], links[i][5], false );
            }
        }
        
        cll.absolute &= ~Result.CALCULATED;
        
    }
    
    // used to recalculate references when deleting/adding rows/columns
    public void shiftReferences( int I1, int J1, int dI, int dJ ) {
        nLinks = 0;			// clear the cross reference table, will create anew
        for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
            Result cell = (Result) cll.nextElement();
            if( cell.isFormula() ) {
                Result op = Result.RESULT_ERROR;        // an internal error if occurs
                try {
                    op = Result.Evaluate( cell.str );
                } catch( Exception ee ) { /* old formula ought to be OK here */ }
                if( op.shiftReferences( (short)I1,(short)J1, (short)dI,(short)dJ ) ) {

                    Result rr = Result.RESULT_ERROR;
                    try {
                        rr = op.calculate( this ); // new value
                    }
                    catch( BadFormulaException e ) { }

                    cell.str = '=' + op.toString(); // replace formula - we may not need it, just in case

                    if( !rr.equals( cell.funcargs[0] ) ) {
                        cell.funcargs[0] = rr;
                        calculateDepended( cell.i1, cell.j1, true );  // bug - dependencies aren't ready yet (were cleared), will be changed in wave algoritm
                                                            // to do - use calculated[][], remember changes and recalculate them at once outside the modification routine
                    }
                }
                addReferences( op, cell.i1, cell.j1 );	// new references
            }
        }
    }
    
    
    public void resize( int newrows, int newcols ) {
        
        // remove cells outside the new size
        for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
            Result cell = (Result) cll.nextElement();
            if( cell.i1 >= newrows || cell.j1 >= newcols )
                cells.remove( cell );
        }
        
        if( newrows > rows || newcols > columns ) {
            
            short _columnWidth[] = new short[newcols];
            short _rowHeight[] = new short[newrows];
            
            for( int i=0; i<newrows; i++ )
                if( i < rows )
                    _rowHeight[i] = rowHeight[i];
                else
                    _rowHeight[i] = (short) defaultHeight;
            
            for( int j=0; j<newcols; j++ )
                if( j < columns )
                    _columnWidth[j] = columnWidth[j];
                else
                    _columnWidth[j] = (short) defaultWidth;
            
            rowHeight = _rowHeight;
            columnWidth = _columnWidth;
            
        }
        rows = newrows;
        columns = newcols;
        changed = true;
    }
    
    // insert rows/columns (no formula modifications)
    public void insertCells( int I1, int J1, int dI, int dJ, boolean resize ) {
        
        if( resize ) resize( rows+dI, columns+dJ );

        // will copy cells into a new hastable
        Hashtable _cells = new Hashtable( cells.size() );

        for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
            Result cell = (Result) cll.nextElement();
            boolean fDelete = false;
            if( cell.i1 >= I1 ) {
                if( (dI > 0 && cell.i1+dI >= rows) || (dI < 0 && cell.i1+dI < I1) )
                    fDelete = true;            // remove it - outside the area or inside cut area
                else
                    cell.i1 += dI;
            }
            if( cell.j1 >= J1 ) {
                if( (dJ > 0 && cell.j1+dJ >= columns) || (dJ < 0 && cell.j1+dJ < J1) )
                    fDelete = true;            // remove it - outside the area or inside cut area
                else
                    cell.j1 += dJ;
            }
            if( fDelete ) {
                // remove from the crossref table
                clear( cell.i1, cell.j1 );
            }
            else {
                _cells.put( cell, cell );
            }
        }
        
        cells = _cells;                     // replace cells
        
        if( dI > 0 ) {
            for( int i=rows-dI-1; i>=I1; i-- ) {
                rowHeight[i+dI] = rowHeight[i];
            }
        }
        else if( dI < 0 ) {
            for( int i=I1; i<rows+dI; i++ ) {
                rowHeight[i] = rowHeight[i-dI];
            }
        }
        
        if( dJ > 0 ) {
            // move
            for( int j=columns-dJ-1; j>=J1; j-- ) {
                columnWidth[j+dJ] = columnWidth[j];
            }
        }
        else if( dJ < 0 ) {
            for( int j=J1; j<columns+dJ; j++ ) {
                columnWidth[j] = columnWidth[j-dJ];
            }
        }
        
        shiftReferences( I1, J1, dI, dJ );
        
        changed = true;
    }
    
    public void copyCell1( int i, int j, int di, int dj ) {
        
        if( di == 0 && dj == 0 ) return;
        
        int newI = i+di;
        int newJ = j+dj;

        Result cell = getCell( i, j );
        Result newcell = new Result( cell );
        
        clear( newI, newJ );            // clear old formula references, if any
        
        newcell.i1 = (short) newI;
        newcell.j1 = (short) newJ;
        
        if( cell.isFormula() ) {
            Result op = Result.RESULT_ERROR;
            try {
                op = Result.Evaluate( cell.str );
            }
            catch( Exception ee ) { /* internal error - formula cannot be wrong */ }
            if( op.moveReferences( di, dj ) ) {         // bug? shouldn't it calculate depended even an exception occurs

                Result rrr;
                try {
                    rrr = op.calculate( this ); // new value
                }
                catch( Exception e ) {
                    rrr = Result.RESULT_ERROR;
                }

                newcell.funcargs = new Result[1];
                newcell.funcargs[0] = rrr;
                newcell.str = '=' + op.toString();

                calculateDepended( newI, newJ, true );  // bug - dependencies aren't ready yet (old values)
                                                        // to do - use calculated[][], remember changes and recalculate them at once outside the modification routine
                                                        // (wave algorithm)
            }
            
            addReferences( op, newI, newJ );	// new references
            
        }
        
        cells.put( newcell, newcell );
        
    }    

    public void copyRow1( int i, int di ) {
        for( int j=0; j<columns; j++ )
            copyCell1( i,j, di,0 );
    }
    
    public void incrementServerVersion() {
        serverVersion = ( (serverVersion >> 16) + 1 ) << 16;
    }

    public int getServerVersion() {
        return serverVersion >> 16;
    }

    // uses proxy pattern
    public Result getCell( int i, int j ) {
        Result rr = (Result) cells.get( new Integer( (i<<16)|j ) );
        if( rr == null )
            rr = Result.createCell( i,j, "", Result.RESULT_EMPTY );
        return rr;
    }
    
    public Result getCellValue( int i, int j ) {
        return getCell( i, j ).funcargs[0];
    }

    /* ========================================================================
     * converting to SYLK format
     */
    private static final String SYLK_HEADER =
	"ID;PMICROCALC;N;E\r\nP;PGeneral\r\nO;L;D;V0;K47;G100 0.001\r\n";
    private static final String SYLK_TAIL =
	"E\r\n";
    private static final long DATE_SHIFT	= 25569L*24*60L*60L*1000L;

    public String toSylk() {

	StringBuffer sb = new StringBuffer(cells.size()*30);
        
        sb.append( SYLK_HEADER );

	sb.append( "B;Y" );
	sb.append( Integer.toString(rows) );
	sb.append( ";X" );
	sb.append( Integer.toString(columns) );
	sb.append( "\r\n" );

        for( Enumeration cll = cells.elements(); cll.hasMoreElements() ; ) {
            Result cell = (Result) cll.nextElement();
		    Result val = cell.funcargs[0];
		    String form = cell.str;
		    Result operand = Result.RESULT_ERROR;
                    try {
                        operand = Result.Evaluate( form );
                    }
                    catch( Exception ee ) { }
		
		    sb.append( "C;Y" );
		    sb.append( Integer.toString( cell.i1+1 ) );
		    sb.append( ";X" );
		    sb.append( Integer.toString( cell.j1+1 ) );
		    sb.append( ';' );
		
		    sb.append( 'K' );
		    if( val.type == Result.TYPE_STRING ) {
			sb.append( '"' );
		    	sb.append( val );
			sb.append( '"' );
		    }
		    else if( val.type == Result.TYPE_DATE ) {
		        sb.append( Result.millisToBcd( val.ll+DATE_SHIFT ) );
		    } else {
		    	sb.append( val );
		    }

		    if( form.charAt(0) == '=' ) {
		    	sb.append( ';' );
			sb.append( 'E' );
			sb.append( val.type == Result.TYPE_ERROR ? form : getSylkFormula( operand, cell.i1, cell.j1 ) );
		    }
		    sb.append( "\r\n" );
		}

        sb.append( SYLK_TAIL ); 

	return sb.toString();
    }

    private static String getSylkAddress( int dsti, int dstj, int srci, int srcj, boolean absi, boolean absj ) {
	StringBuffer sb = new StringBuffer(10);
	int di = dsti - srci;
	int dj = dstj - srcj;
	sb.append( 'R' );
	if( absi ) {
	    sb.append( dsti+1 );
	}
	else if( di != 0 ) {
	    sb.append( '[' );
	    sb.append( Integer.toString( di ) );
	    sb.append( ']' );
	}
	sb.append( 'C' );
	if( absj ) {
	    sb.append( dstj+1 );
	}
	else if( dj != 0 ) {
	    sb.append( '[' );
	    sb.append( Integer.toString( dj ) );
	    sb.append( ']' );
	}
	return sb.toString();
    }

    // gets formula in the SYLK format, recursive
    private static String getSylkFormula( Result rr, int i, int j ) {
	String ss="";
	String s1=null, s2=null;
        if( rr.funcargs.length > 0 ) s1 = getSylkFormula( rr.funcargs[0], i, j );
        if( rr.funcargs.length > 1 ) s2 = getSylkFormula( rr.funcargs[1], i, j );
	switch( rr.type ) {
	    case Result.TYPE_BOOLEAN:
		ss = Result.booleanToString( rr.ll );
		break;
	    case Result.TYPE_DATE:
		ss = '"' + rr.toString() + '"';
		break;
	    case Result.TYPE_LONG:
	    	ss = Long.toString( rr.ll );
		break;
	    case Result.TYPE_BCD:
		ss = rr.toString();
		break;
	    case Result.TYPE_STRING:
		ss = '"' + rr.str + '"';
		break;
	    case Result.TYPE_RANGE:
		ss = getSylkAddress( rr.i1, rr.j1, i, j, (rr.absolute&Result.ABS_I1)!=0,(rr.absolute&Result.ABS_J1)!=0 );
	    	if( rr.i1 != rr.i2 || rr.j1 != rr.j2 )
		    ss += ':' + getSylkAddress( rr.i2, rr.j2, i, j, (rr.absolute&Result.ABS_I2)!=0,(rr.absolute&Result.ABS_J2)!=0 );
		break;
	    case Result.OPER_UMIN:
		ss = '-' + s1;
		break;
	    case Result.OPER_ADD:
		ss = s1 + '+' + s2;
		break;
	    case Result.OPER_SUB:
		ss = s1 + '-' + s2;
		break;
	    case Result.OPER_MUL:
		ss = s1 + '*' + s2;
		break;
	    case Result.OPER_DIV:
		ss = s1 + '/' + s2;
		break;
//	    case Result.OPER_POW:
//		ss = s1 + '^' + s2;
//		break;
	    case Result.OPER_CONCAT:
		ss = s1 + '&' + s2;
		break;
	    case Result.OPER_EQ:
		ss = s1 + '=' + s2;
		break;
	    case Result.OPER_NE:
		ss = s1 + "!=" + s2;
		break;
	    case Result.OPER_LT:
		ss = s1 + '<' + s2;
		break;
	    case Result.OPER_GT:
		ss = s1 + '>' + s2;
		break;
	    case Result.OPER_LE:
		ss = s1 + "<=" + s2;
		break;
	    case Result.OPER_GE:
		ss = s1 + ">=" + s2;
		break;
	    case Result.OPER_BRACK:
		ss = '(' + s1 + ')';
		break;
	    case Result.TYPE_FUNC:
		ss = rr.getFunctionName() + '(';
		for( int ii=0; ii<rr.funcargs.length; ii++ ) {
		    ss += getSylkFormula( rr.funcargs[ii], i, j );
		    if( ii!=rr.funcargs.length-1 ) ss += ',';
		}
		ss += ')';
		break;
	    case Result.TYPE_ERROR:					// TYPE_ERROR
		ss = "#ERR";
		break;
	    default:
		System.out.println("internal error: unknow Result type");
	}
	return ss;
    }


}

