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
package com.wapindustrial.calc;

import net.jscience.math.MathFP;

import java.util.*;
import java.io.*;

public final class Result {
    
    // constants
    
    // types
    public static final byte TYPE_EMPTY	= 0;
    public static final byte TYPE_BOOLEAN=1;
    public static final byte TYPE_DATE	= 2;
    public static final byte TYPE_LONG	= 3;
    public static final byte TYPE_BCD	= 4;
    public static final byte TYPE_STRING= 5;
    
    public static final byte TYPE_RANGE	= 15;
    
    // operators
    public static final byte OPER_UMIN	= 20;
    public static final byte OPER_ADD	= 21;
    public static final byte OPER_SUB	= 22;
    public static final byte OPER_MUL	= 23;
    public static final byte OPER_DIV	= 24;
    public static final byte OPER_POW	= 25;
    public static final byte OPER_CONCAT= 26;
    public static final byte OPER_EQ	= 27;
    public static final byte OPER_NE	= 28;
    public static final byte OPER_LT	= 29;
    public static final byte OPER_GT	= 30;
    public static final byte OPER_LE	= 31;
    public static final byte OPER_GE	= 32;
    public static final byte OPER_BRACK	= 33;
    
    public static final byte TYPE_FUNC	= 50;
    public static final byte TYPE_CELL	= 51;

    public static final byte TYPE_ERROR	= 100;
    
    // functions (stored in ll field)
    private static final byte FUNC_SUM	= 0;
    private static final byte FUNC_IF	= 1;
    private static final byte FUNC_AND	= 2;
    private static final byte FUNC_OR	= 3;
    private static final byte FUNC_POW	= 4;
    private static final byte FUNC_EXP	= 5;
    private static final byte FUNC_LN	= 6;
    private static final byte FUNC_SQRT	= 7;
    private static final byte FUNC_SIN	= 8;
    private static final byte FUNC_COS	= 9;
    private static final byte FUNC_TAN	= 10;
    private static final byte FUNC_COT	= 11;
    private static final byte FUNC_ASIN	= 12;
    private static final byte FUNC_ACOS	= 13;
    private static final byte FUNC_ATAN	= 14;
    private static final byte FUNC_ATAN2= 15;
    private static final byte FUNC_ABS	= 16;
    private static final byte FUNC_PI	= 17;
    private static final byte FUNC_E	= 18;
    
    private static final String FUNCTION_NAMES[] = { 
        "SUM",
        "IF",
        "AND",
        "OR",
        "POW",
        "EXP",
        "LN",
        "SQRT",
        "SIN",
        "COS",
        "TAN",
        "COT",
        "ASIN",
        "ACOS",
        "ATAN",
        "ATAN2",
        "ABS",
        "PI",
        "E"
    };

    // error messages
    private static final String MSG_INTERNAL_ERROR      = "Internal error";
    private static final String MSG_ILLEGAL_DATE_OPER   = "illegal date operation";
    private static final String MSG_WRONG_ARG_NUMBER    = "wrong number of arguments of function";
    private static final String MSG_WRONG_ARG_TYPE      = "wrong type of function argument";
    
    
    // format masks
/*    
    public static final short FORMAT_JUSTIFY_LEFT	= 0x0001;
    public static final short FORMAT_JUSTIFY_RIGHT	= 0x0002;
    public static final short FORMAT_JUSTIFY_CENTER	= 0x0003;
    public static final short FORMAT_JUSTIFY_TOP	= 0x0004;
    public static final short FORMAT_JUSTIFY_BOTTOM	= 0x0008;
    public static final short FORMAT_JUSTIFY_VCENTER	= 0x000C;
  */  
    public static final short FORMAT_JUSTIFY_MASK	= 0x000F;

    public static final short FORMAT_NUMBER	= 0x0100;	// next byte - numbers after decimal dot
    public static final short FORMAT_NUMBER_SEP	= 0x1000;	// thousand separator present
    public static final short FORMAT_DATE	= 0x0200;	//
    public static final short FORMAT_DATE_TIME	= 0x0010;	//
    public static final short FORMAT_DATE_DATE	= 0x0020;	// dd/mm/yy
    public static final short FORMAT_DATE_SECONDS= 0x0040;	// dd/mm/yy hh:mm
    public static final short FORMAT_DATE_YEAR4	= 0x0080;	// hh:mm
    public static final short FORMAT_DATE_ORDER	= 0x0400;	// mm first
    public static final short FORMAT_DATE_SEP	= 0x0800;	// dot separator ( dd.mm.yyyy )
    public static final short FORMAT_TYPE_MASK	= 0x1FF0;
    
    // bits of absolute (should be named state after all)
    public static final byte ABS_I1	= 0x01;             // are used in range
    public static final byte ABS_J1	= 0x02;
    public static final byte ABS_I2	= 0x04;
    public static final byte ABS_J2	= 0x08;
    public static final byte CALCULATED	= 0x10;             // temporary, used in cell datatype
    public static final byte FORMULA	= 0x20;             // is set if it's a formula expression (beginning with '=')
    
    private static final String EMPTY_FORMULA  = "";
    
    private static final long DAY_BCD_MS_LONG = 24L*60L*60L*1000L;
    private static final long DAY_BCD_MS = MathFP.toFP( DAY_BCD_MS_LONG );
    
    private static final Result FUNCARGS_ZERO[] = new Result[0];
    
    public static final Result RESULT_ERROR = new Result( TYPE_ERROR );
    public static final Result RESULT_EMPTY = new Result( TYPE_EMPTY );
    
    public byte type;
    public long ll;				// keeps long & date (ms) & bcd, and boolean (0 and not 0)
                                                // TYPE_CELL: stores formats
    public String str = EMPTY_FORMULA;          // string constant & function name, and formula for cells (cannot be null)
    public short i1,i2,j1,j2;			// range, for cell (i1,j1) keep its address
    public byte absolute=0;			// absolute flags for range, and calculated flag
    public Result funcargs[] = FUNCARGS_ZERO;   // for cell funcargs[0] keeps its value
    
    // to parse the formula
    private static char buf[];
    private static int nBuf;
    
    private static boolean inExpression;	// true when toString() is called in an expression
    private static StringBuffer sBuf;
    
    // ===============================================
    // constructors
    //
    public Result() {
        type = TYPE_EMPTY;
    }
    
    public Result( byte type ) {
        this.type = type;
    }
    
    public Result( byte _type, long ll ) {
        type = _type;
        this.ll = ll;
    }
    
    public Result( long ll ) {
        type = TYPE_LONG;
        this.ll = ll;
    }
    
    public Result( String str ) {
        type = TYPE_STRING;
        this.str = str;
    }
    
    Result( boolean _bool ) {
        type = TYPE_BOOLEAN;
        ll = _bool ? 1 : 0;
    };

    public Result( short _i1, short _j1, short _i2, short _j2, byte _absolute ) {
        type = TYPE_RANGE;
        i1 = _i1; j1 = _j1; i2 = _i2; j2 = _j2;
	absolute = _absolute;
        normalizeRange();
    }
    
    public Result( byte _type, Result _oper1, Result _oper2 ) {
        type = _type;
        if( _oper2 != null ) {
            funcargs = new Result[2];
            funcargs[0] = _oper1;
            funcargs[1] = _oper2;
        }
        else {
            funcargs = new Result[1];
            funcargs[0] = _oper1;
        }
    }
    
    public Result( byte func, Result _funcargs[] ) {
        type = TYPE_FUNC;
        ll = (long) func;
        funcargs = _funcargs;
    }

    // does not clone bcd, str, and opers
    public Result( Result rr /*, boolean clone*/ ) {
        if( rr == null )			// empty
            return;
        type = rr.type;
        ll = rr.ll;
        i1 = rr.i1;
        j1 = rr.j1;
        i2 = rr.i2;
        j2 = rr.j2;
        absolute = rr.absolute;
        str = rr.str;                   // shouldn't be cloned?
/*        if( clone ) {
            funcargs = new Result[rr.funcargs.length];
            for( int i=0; i<rr.funcargs.length; i++ )
                funcargs[i] = rr.funcargs[i];           // don't need to clone items since it's actual only for cell values at the moment
        }
        else {*/
            funcargs = rr.funcargs;
//        }
                
    }
/*
    public static Result clone( Result rr ) {
        if( rr == null )			// empty
            return null;
        Result r = new Result( rr );
        r.funcargs = new Result[rr.funcargs.length];
        for( int i=0; i<rr.funcargs.length; i++ )
            r.funcargs[i] = clone( rr.funcargs[i] );
        return r;
    }
  */  
    // only for constants (not expressions)
    /** Reads an Result from input stream
     * @param dinp the stream to read from
     * @see #save
     * @throws BadFormulaException when can't create a BCD datatype from byte string
     * @throws IOException when can't read data from input stream
     */    
    public Result( DataInputStream dinp, int sheetVersion ) throws BadFormulaException,IOException {
        type = dinp.readByte();
        switch( type ) {
            
            case TYPE_EMPTY:
            case TYPE_ERROR:
                break;
                
            case TYPE_BOOLEAN:
                ll = dinp.readByte();
                break;
            case TYPE_LONG:
            case TYPE_DATE:
            case TYPE_BCD:
                ll = dinp.readLong();
                break;
            case TYPE_STRING:
                str = dinp.readUTF();
                break;
/*            case TYPE_CELL:
                i1 = dinp.readShort();
                j1 = dinp.readShort();
                funcargs = new Result[1];
                funcargs[0] = new Result( dinp, sheetVersion );     // value
                ll = (long) dinp.readShort();                       // format
                str = dinp.readUTF();                               // formula
                if( str.length() > 0 && str.charAt(0) == '=' )
                    absolute |= FORMULA;
                break;
 */
            default:
                throw new BadFormulaException(MSG_INTERNAL_ERROR);
        }
    }

    public static Result restoreCell( DataInputStream dinp, int sheetVersion ) throws BadFormulaException,IOException {
        int i1 = dinp.readShort();
        int j1 = dinp.readShort();
        Result val = new Result( dinp, sheetVersion );     // value
        long ll = (long) dinp.readShort();                       // format
        String str = dinp.readUTF();                               // formula
        Result rr = createCell( i1,j1, str, val );
        rr.ll = ll;
        return rr;
    }
    
    /** stores Result to the output stream
     * only for constants (not expressions)
     * @param dout Stream to save
     * @throws IOException if an DataOutputStream error occurs, or it's not a constant
     */    
    public void save( DataOutputStream dout, boolean writeType ) throws IOException {
        if( writeType ) dout.writeByte( type );
        switch( type ) {
            case TYPE_EMPTY:
            case TYPE_ERROR:
                break;
            case TYPE_BOOLEAN:
                dout.writeByte( (byte) ll );
                break;
            case TYPE_LONG:
            case TYPE_DATE:
            case TYPE_BCD:
                dout.writeLong( ll );
                break;
            case TYPE_STRING:
                dout.writeUTF( str );
                break;
            case TYPE_CELL:
                dout.writeShort( i1 );
                dout.writeShort( j1 );
                funcargs[0].save( dout, true );                           // value
                dout.writeShort( (short) ll );                      // format
                dout.writeUTF( str );                               // formula
                break;
            default:
                throw new IOException( MSG_INTERNAL_ERROR );
        }
    }

    public static String removeZeros( String ss ) {
        StringBuffer currVal = new StringBuffer( ss );
	char cc;
	int i;
        for( i = currVal.length()-1; (cc=currVal.charAt(i)) == '0'; i--)
            currVal.setLength( i );

        if(cc == '.')
            currVal.setLength( i );
        
        return currVal.toString();
    }
    
    public String toString() {
        return toString( false,0 );
    }
    
    public String toString( boolean fExpr, int formats ) {
        String ss="";
        String s1=null, s2=null;
        if( funcargs.length > 0 ) s1 = funcargs[0].toString(true,0);
        if( funcargs.length > 1 ) s2 = funcargs[1].toString(true,0);
        switch( type ) {
            case TYPE_EMPTY:
                ss = "";
                break;
            case TYPE_BOOLEAN:
                ss = booleanToString( ll );
                break;
            case TYPE_DATE:
                ss = dateToString( ll, formats );
                if( fExpr )
                    ss = '#' + ss + '#';
                break;
            case TYPE_LONG:
                ss = Long.toString( ll ) + "L";
                break;
            case TYPE_BCD:
                if( (formats&FORMAT_NUMBER) != 0 ) {
                    int nn = (formats & 0x00F0) >> 4;
                    ss = removeZeros( MathFP.toString( ll, nn ) );
                }
                else
                    ss = removeZeros( MathFP.toString( ll, MicroCalc.DEFAULT_PRECISION ) );
                break;
            case TYPE_STRING:
                ss = str;
                if( fExpr ) ss = '"' + ss + '"';
                break;
            case TYPE_RANGE:
                ss = rangeAddress( i1,j1,i2,j2, absolute );
                break;
            case OPER_UMIN:
                ss = '-' + s1;
                break;
            // todo - this may be reduced by using array of oper names, and by StringBuffer    
            // shouldn't we declare oper class, really?
            case OPER_ADD:
                ss = s1 + '+' + s2;
                break;
            case OPER_SUB:
                ss = s1 + '-' + s2;
                break;
            case OPER_MUL:
                ss = s1 + '*' + s2;
                break;
            case OPER_DIV:
                ss = s1 + '/' + s2;
                break;
            case OPER_POW:
                ss = s1 + '^' + s2;
                break;
            case OPER_CONCAT:
                ss = s1 + '&' + s2;
                break;
            case OPER_EQ:
                ss = s1 + '=' + s2;
                break;
            case OPER_NE:
                ss = s1 + "!=" + s2;
                break;
            case OPER_LT:
                ss = s1 + '<' + s2;
                break;
            case OPER_GT:
                ss = s1 + '>' + s2;
                break;
            case OPER_LE:
                ss = s1 + "<=" + s2;
                break;
            case OPER_GE:
                ss = s1 + ">=" + s2;
                break;
            case OPER_BRACK:
                ss = '(' + s1 + ')';
                break;
            case TYPE_FUNC:
                ss = FUNCTION_NAMES[(int)ll] + '(';
                for( int i=0; i<funcargs.length; i++ ) {
                    ss += funcargs[i].toString(true,0);
                    if( i!=funcargs.length-1 ) ss += ',';
                }
                ss += ')';
                break;
            case TYPE_ERROR:					// TYPE_ERROR
                ss = "#ERR";
        }
        return ss;
    }
    
    private void normalizeRange() {
        if( i1 > i2 ) { short ii=i1; i1=i2; i2=ii; absolute = (byte) ( (absolute&~(ABS_I1|ABS_I2)) | (absolute&ABS_I1<<2) | (absolute&ABS_I2>>2) ); }
        if( j1 > j2 ) { short jj=j1; j1=j2; j2=jj; absolute = (byte) ( (absolute&~(ABS_J1|ABS_J2)) | (absolute&ABS_J1<<2) | (absolute&ABS_J2>>2) ); }
        if( i1 == i2 ) absolute = (byte)( (absolute&~ABS_I2) | ((absolute&ABS_I1)<<2) );
        if( j1 == j2 ) absolute = (byte)( (absolute&~ABS_J2) | ((absolute&ABS_J1)<<2) );
    }
    
    public static String columnName( int j ) {
        StringBuffer name = new StringBuffer(2);
        final int nn = 'Z' - 'A' + 1;
        int n1 = j / nn;
        if( n1 > 0 ) name.append( (char)('A'+n1-1) );
        j %= nn;
        name.append( (char)('A'+j) );
        return name.toString();
    }
    
    public static String cellAddress( int i, int j, int absolute ) {
	StringBuffer sb = new StringBuffer( 10 );
	if( (absolute&ABS_J1) != 0 ) sb.append( '$' );
        sb.append( columnName( j ) );
	if( (absolute&ABS_I1) != 0 ) sb.append( '$' );
	sb.append( String.valueOf( i+1 ) );
        return sb.toString();
    }
    
    public static String rangeAddress( int i1, int j1, int i2, int j2, int absolute ) {
        StringBuffer sb = new StringBuffer( cellAddress( i1,j1, absolute ) );
        if( i1 != i2 || j1 != j2 ) {
            sb.append( ':' );
            sb.append( cellAddress( i2,j2, absolute>>2 ) );
        }
        return sb.toString();
    }
    
    public static Result parseAddress( String ss ) throws BadFormulaException {
        byte sss[] = ss.toUpperCase().getBytes();
        int ii = 0;
        char cc;
        try {
	    byte abs = 0;
            final int nn = 'Z' - 'A';
            short col = 0;
            short row = 0;
	    if( sss[0] == '$' ) {
		ii = 1;
		abs = ABS_J1;
	    }
            while( !( Character.isDigit( cc=(char)(sss[ii]) ) || cc == '$' ) ) {
                if( cc < 'A' || cc > 'Z' )
                    throw new Exception();
                col *= nn;
                col += (short)(cc - 'A');
                ii++;
                if( col >= nn*nn ) throw new Exception();	// only 2 chars in column name are allowed
            }
            if( ii >= sss.length )
                throw new Exception();
	    if( cc == '$' ) {
		ii++;
		abs |= ABS_I1;
	    }
            while( ii < sss.length && Character.isDigit( cc=(char)(sss[ii]) ) ) {
                row *= 10;
                row += (short)(cc - '0');
                ii++;
            }
            row--;
            if( row < 0 )
                throw new Exception();
            if( ii == sss.length )
                return new Result( row,col,row,col, abs );
            if( cc == ':' ) {
                //		ii++;
                if( ss.indexOf( ':', ii+1 ) != -1 )
                    throw new Exception();		// more then 1 occurency of ':'
                Result rr = parseAddress( ss.substring( ii+1 ) );
                rr.i1 = row; rr.j1 = col;
		rr.absolute = (byte) ( (rr.absolute << 2) | abs );
                rr.normalizeRange();
                return rr;
            }
            if( ii != sss.length )
                throw new Exception();		// there are characters after digits
        }
        catch( Exception e ) { }
        throw new BadFormulaException( "illegal cell address [" + ss + "]" );
    }
    
    private Result sumRange( Sheet sheet ) throws BadFormulaException {
        //	if( type != TYPE_RANGE )
        //	    return this;
        Result rr = new Result( TYPE_BCD );             // zero
        Result val;
        for( int i=i1; i<=i2; i++ )
            for( int j=j1; j<=j2; j++ ) {
                val = sheet.getCellValue( i,j );
                if( val.type == TYPE_EMPTY || val.type == TYPE_STRING ) continue;
                rr = operation( OPER_ADD, rr, val );
            }
        return rr;
    }
    
    private Result toMaxType( int t2 ) throws BadFormulaException {
        
        // TYPE_RANGE is illegal, call calculate() before this
        
        if( type < t2 ) {
            if( t2 == TYPE_STRING ) {
                String ss="";			// TYPE_EMPTY
                if( type == TYPE_BOOLEAN )	ss = booleanToString( ll );
                if( type == TYPE_DATE )		ss = dateToString( ll,0 );
                if( type == TYPE_LONG )		ss = Long.toString( ll );
                if( type == TYPE_BCD )		ss = removeZeros( MathFP.toString( ll ) );
                return new Result( ss );
            }
            if( t2 == TYPE_BCD ) {
                long bcd = 0L;	// TYPE_EMPTY
                if( type == TYPE_BOOLEAN )	if( ll!=0 ) bcd = MathFP.ONE;
                if( type == TYPE_DATE )		bcd = MathFP.div( ll, DAY_BCD_MS_LONG );
                if( type == TYPE_LONG )		bcd = MathFP.toFP( ll );
                return createFloat( bcd );
            }
            if( t2 == TYPE_LONG ) {
                long lll = 0L;			// TYPE_EMPTY
                if( type == TYPE_BOOLEAN || type == TYPE_DATE )	lll = ll;
                return new Result( lll );
            }
            if( t2 == TYPE_DATE ) {
                long lll = 0L;			// TYPE_EMPTY
                if( type == TYPE_BOOLEAN )
                    throw new BadFormulaException("cannot convert boolean to datetime");
                return new Result( TYPE_DATE, lll );
            }
            if( t2 == TYPE_BOOLEAN ) {
                if( type == TYPE_EMPTY ) return new Result( false );	// zero
            }
            if( t2 == TYPE_ERROR ) {
                return RESULT_ERROR;	// error
            }
        }
        return this;				// type >= t2
    }
    
    public static String booleanToString( long bb ) { return bb!=0 ? "TRUE" : "FALSE"; }
    
    public static String dateToString( long ms, int formats ) {
        StringBuffer sb = new StringBuffer(20);
        
        Calendar cd = Calendar.getInstance();
        cd.setTime( new Date( ms ) );
        
        int year = cd.get( Calendar.YEAR );
        int day = cd.get( Calendar.DAY_OF_MONTH );
        int month = cd.get( Calendar.MONTH )+1;
        int hour = cd.get( Calendar.HOUR_OF_DAY );
        int min = cd.get( Calendar.MINUTE );
        int sec = cd.get( Calendar.SECOND );
        
        if( formats == 0 ) formats = FORMAT_DATE_DATE;
        
        if( (formats&FORMAT_DATE_DATE) != 0 ) {
            char sep = (formats&FORMAT_DATE_SEP)!=0 ? '.' : '/';
            if( (formats&FORMAT_DATE_ORDER) == FORMAT_DATE_ORDER )
                sb.append( Integer.toString( month ) );
            else
                sb.append( Integer.toString( day ) );
            sb.append( sep );
            if( (formats&FORMAT_DATE_ORDER)!=0 ) {
                if( day < 10 ) sb.append( '0' );
                sb.append( Integer.toString( day ) );
            }
            else {
                if( month < 10 ) sb.append( '0' );
                sb.append( Integer.toString( month ) );
            }
            sb.append( sep );
            String ss = Integer.toString( year );
            if( (formats&FORMAT_DATE_YEAR4) == 0 )
                ss = ss.substring(2);
            sb.append( ss );
        }
        
        if( (formats&FORMAT_DATE_TIME)!=0 ) {
            if( (formats&FORMAT_DATE_DATE) != 0 )
                sb.append( ' ' );
            if( hour < 10 ) sb.append( '0' );
            sb.append( Integer.toString( hour ) );
            sb.append( ':' );
            if( min < 10 ) sb.append( '0' );
            sb.append( Integer.toString( min ) );
        }
        
        return sb.toString();
    }
    
    public static long bcdToMillis( long bcd ) {
//        return MathFP.mul( bcd, DAY_BCD_MS_LONG ).toLong();
        return MathFP.toLong( MathFP.mul( bcd, DAY_BCD_MS ) );
    }
    
    public static long millisToBcd( long ll ) {
        return MathFP.div( MathFP.toFP( ll ), DAY_BCD_MS_LONG );
    }
    
    public static Result operation( int oper, Result r1, Result r2 ) throws BadFormulaException {
        
        if( !r1.isConstant() || !r2.isConstant() ) throw new BadFormulaException( MSG_INTERNAL_ERROR );
        
        // test illegal combinations for dates
        boolean d1 = r1.type == TYPE_DATE;
        boolean d2 = r2.type == TYPE_DATE;
        if( d1 || d2 ) {
            if( d1 && d2 ) {
                if( oper != OPER_SUB && !(oper >= OPER_EQ && oper <= OPER_GE) )
                    throw new BadFormulaException(MSG_ILLEGAL_DATE_OPER);
            }
            else {
                
                if( oper != OPER_SUB && oper != OPER_ADD )
                    throw new BadFormulaException(MSG_ILLEGAL_DATE_OPER);
                // one of the args isn't date
                // swap args, date the first
                Result rr = d1 ? r1 : r2;
                r2 = d1 ? r2 : r1;
                r1 = rr;
                if( r2.type != TYPE_STRING ) {	// string in the normal way by converting date to string
                    
                    if( r2.type == TYPE_LONG )	// milliseconds
                        return new Result( TYPE_DATE, oper == OPER_ADD ? r1.ll + r2.ll : r1.ll - r2.ll );
                        
                        if( r2.type == TYPE_BCD ) {	// days
                            long lll = bcdToMillis( r2.ll );
                            lll = oper == OPER_ADD ? r1.ll + lll : r1.ll - lll;
                            Result _rr = new Result( TYPE_DATE, lll );
                            return _rr;
                        }
                        
                        if( r2.type == TYPE_EMPTY )
                            return new Result( r1 );
                        
                        throw new BadFormulaException(MSG_ILLEGAL_DATE_OPER);	// +/- with other datatimes than LONG & BCD & STRING
                }
            }
            
        }
        
        if( oper == OPER_CONCAT ) {
            r1 = r1.toMaxType( TYPE_STRING );
            r2 = r2.toMaxType( TYPE_STRING );
        }
        // to max type of both
        r1 = r1.toMaxType( r2.type );
        r2 = r2.toMaxType( r1.type );
        int tp = r1.type;
        
        if( tp == TYPE_ERROR || tp == TYPE_EMPTY )
            return r1;
        
        switch( oper ) {
            case OPER_EQ:
                return new Result( r1.equals( r2 ) );	// very ineffective for most datatypes, should be separated
            case OPER_NE:
                return new Result( !r1.equals( r2 ) );	// the same
            case OPER_ADD:
                if( tp == TYPE_LONG || tp == TYPE_BOOLEAN ) return new Result( r1.ll + r2.ll );
                if( tp == TYPE_BCD ) return createFloat( r1.ll + r2.ll );
                if( tp == TYPE_STRING )	return new Result( r1.str + r2.str );
            case OPER_SUB:
                if( tp == TYPE_DATE ) {
                        return createFloat( MathFP.div( MathFP.toFP(r1.ll - r2.ll), DAY_BCD_MS ) );
                }
                if( tp == TYPE_LONG || tp == TYPE_BOOLEAN ) return new Result( r1.ll - r2.ll );
                if( tp == TYPE_BCD )	return createFloat( r1.ll - r2.ll );
                if( tp == TYPE_STRING )	return RESULT_ERROR;
            case OPER_MUL:
                if( tp == TYPE_LONG || tp == TYPE_BOOLEAN ) return new Result( r1.ll * r2.ll );
                if( tp == TYPE_BCD )	return createFloat( MathFP.mul( r1.ll, r2.ll ) );
                if( tp == TYPE_STRING )	return RESULT_ERROR;
            case OPER_DIV:
                // needs more checks for DIV/0
                if( tp == TYPE_BOOLEAN )    return RESULT_ERROR;
                if( tp == TYPE_LONG )       return new Result( r1.ll / r2.ll );
                if( tp == TYPE_BCD )        return createFloat( MathFP.div( r1.ll, r2.ll ) );
                if( tp == TYPE_STRING )     return RESULT_ERROR;
            case OPER_POW:
                if( tp == TYPE_BOOLEAN )    return RESULT_ERROR;
                if( tp == TYPE_LONG )       return new Result( MathFP.toLong( MathFP.pow( MathFP.toFP(r1.ll), MathFP.toFP(r2.ll) ) ) );
                if( tp == TYPE_BCD )        return new Result( TYPE_BCD, MathFP.pow( r1.ll, r2.ll ) );
                if( tp == TYPE_STRING )     return RESULT_ERROR;
            case OPER_CONCAT:
                return new Result( r1.str + r2.str );
            case OPER_LE:
                if( r1.equals( r2 ) ) return new Result( true );
                if( tp == TYPE_BOOLEAN ) return new Result( r1.ll == 0L );
                if( tp == TYPE_LONG || tp == TYPE_DATE ) return new Result( r1.ll <= r2.ll );
                if( tp == TYPE_BCD ) return new Result( r1.ll <= r2.ll );
                if( tp == TYPE_STRING ) return new Result( r1.str.length() < r2.str.length() );
            case OPER_LT:
                if( tp == TYPE_BOOLEAN ) return new Result( r1.ll==0L && r2.ll!=0L );
                if( tp == TYPE_LONG || tp == TYPE_DATE ) return new Result( r1.ll < r2.ll );
                if( tp == TYPE_BCD ) return new Result( r1.ll < r2.ll );
                if( tp == TYPE_STRING ) return new Result( r1.str.length() < r2.str.length() );
            case OPER_GE:
                if( r1.equals( r2 ) ) return new Result( true );
                if( tp == TYPE_BOOLEAN ) return new Result( r1.ll!=0L );
                if( tp == TYPE_LONG || tp == TYPE_DATE ) return new Result( r1.ll >= r2.ll );
                if( tp == TYPE_BCD ) return new Result( r1.ll >= r2.ll );
                if( tp == TYPE_STRING ) return new Result( r1.str.length() > r2.str.length() );
            case OPER_GT:
                if( tp == TYPE_BOOLEAN ) return new Result( r1.ll!=0L && r2.ll==0L );
                if( tp == TYPE_LONG || tp == TYPE_DATE ) return new Result( r1.ll > r2.ll );
                if( tp == TYPE_BCD ) return new Result( r1.ll > r2.ll );
                if( tp == TYPE_STRING ) return new Result( r1.str.length() > r2.str.length() );
            default:
                throw new BadFormulaException( MSG_INTERNAL_ERROR );
        }
        
        // unknown operation ?!
        //	return new Result( TYPE_ERROR );
    }
    
    public boolean equals( Result x ) {
        return toString().compareTo( x.toString() ) == 0 ? true : false;
    }
    
    public boolean isOperator() {
        return type >= OPER_UMIN && type <= OPER_GE;
    }
    public boolean isConstant() {
        return (type >= TYPE_EMPTY && type <= TYPE_STRING) || type == TYPE_ERROR;
    }

    public Result calculate( Sheet sheet ) throws BadFormulaException {	// sheet for ranges
        if( isConstant() ) return this;
        if( type == TYPE_RANGE ) {
            if( i1 >= sheet.rows || j2 >= sheet.rows ||
            j1 >= sheet.columns || j2 >= sheet.columns )
                throw new BadFormulaException("reference outside the sheet");
            return sheet.getCell(i1,j1).funcargs[0]; //sheet.values[i1][j1];
        }
        Result a1=null, a2=null;
        if( funcargs.length > 0 ) a1 = funcargs[0].calculate( sheet );
        if( funcargs.length > 1 ) a2 = funcargs[1].calculate( sheet );
        if( type == OPER_BRACK )
            return a1;
        if( type == OPER_UMIN ) {
            switch( a1.type ) {
                case TYPE_LONG:
                    return new Result( TYPE_LONG, -a1.ll );
                case TYPE_BCD:
                    return createFloat( -a1.ll );
            }
            throw new BadFormulaException("unary minus with unsupported data type");
        }
        if( isOperator() )
            return operation( type, a1, a2 );
        if( type == TYPE_FUNC ) {
          return calculateFunc( sheet );
        }
        throw new BadFormulaException( MSG_INTERNAL_ERROR );
    }
    
    private Result calculateFunc( Sheet sheet ) throws BadFormulaException {
      Result a1=null, a2=null;
      int functype = (int) ll;
      switch( functype ) {
          
          case FUNC_E:    
              return createFloat( MathFP.E );
          case FUNC_PI:    
              return createFloat( MathFP.PI );

          case FUNC_SUM:    
            Result y = new Result(TYPE_EMPTY);
            for( int i=0; i<funcargs.length; i++ ) {
                Result x = funcargs[i];
                if( x.type == TYPE_RANGE ) x = x.sumRange( sheet );
                else x = x.calculate(sheet);
                y = operation( OPER_ADD, y, x );
            }
            return y;
          case FUNC_IF:
            if( funcargs.length != 3 )
                throw new BadFormulaException(MSG_WRONG_ARG_NUMBER);
            a1 = funcargs[0].calculate(sheet).toMaxType( TYPE_BOOLEAN );
            a2 = funcargs[1].calculate(sheet);
            Result a3 = funcargs[2].calculate(sheet);
            if( a1.type != TYPE_BOOLEAN )
                throw new BadFormulaException(MSG_WRONG_ARG_TYPE);
            return a1.ll!=0L ? a2 : a3;

          case FUNC_POW:
          case FUNC_ATAN2:
            if( funcargs.length != 2 )
                throw new BadFormulaException(MSG_WRONG_ARG_NUMBER);
            a1 = funcargs[0].calculate(sheet).toMaxType( TYPE_BCD );
            a2 = funcargs[1].calculate(sheet).toMaxType( TYPE_BCD );
            long bcd = functype == FUNC_POW ? MathFP.pow( a1.ll, a2.ll ) : MathFP.atan2( a1.ll, a2.ll );
            return createFloat( bcd );

          // arifm with 1 arg
          case FUNC_SIN:
          case FUNC_COS:
          case FUNC_TAN:
          case FUNC_COT:
          case FUNC_EXP:
          case FUNC_LN:
          case FUNC_SQRT:
          case FUNC_ABS:
          case FUNC_ASIN:
          case FUNC_ACOS:
          case FUNC_ATAN:
            if( funcargs.length != 1 )
                throw new BadFormulaException(MSG_WRONG_ARG_NUMBER);
            a1 = funcargs[0].calculate(sheet).toMaxType( TYPE_BCD );
            if( a1.type != TYPE_BCD )
                throw new BadFormulaException(MSG_WRONG_ARG_TYPE);
            long lll = a1.ll;
            switch( functype ) {
              case FUNC_SIN:
                  return createFloat( MathFP.sin( lll ) );
              case FUNC_COS:
                  return createFloat( MathFP.cos( lll ) );
              case FUNC_TAN:
                  return createFloat( MathFP.tan( lll ) );
              case FUNC_COT:
                  return createFloat( MathFP.cot( lll ) );
              case FUNC_EXP:
                  return new Result( TYPE_BCD, MathFP.exp( lll ) );
              case FUNC_LN:
                  return new Result( TYPE_BCD, MathFP.log( lll ) );
              case FUNC_SQRT:
                  return createFloat( MathFP.sqrt( lll ) );
              case FUNC_ABS:
                  return new Result( TYPE_BCD, lll<0L?-lll:lll );
              case FUNC_ASIN:
                  return createFloat( MathFP.asin( lll ) );
              case FUNC_ACOS:
                  return createFloat( MathFP.acos( lll ) );
              case FUNC_ATAN:
                  return createFloat( MathFP.atan( lll ) );
            } 
            break;
 
          case FUNC_AND:
          case FUNC_OR:
            boolean b1 = functype == FUNC_AND;
            if( funcargs.length != 2 )
                throw new BadFormulaException(MSG_WRONG_ARG_NUMBER);
            boolean rr = b1 ? true : false;
            for( int i=0; i<funcargs.length; i++ ) {
                Result x = funcargs[i].calculate(sheet).toMaxType( TYPE_BOOLEAN );
                if( x.type != TYPE_BOOLEAN )
                    throw new BadFormulaException(MSG_WRONG_ARG_TYPE);
                if( b1 ) rr = rr && x.ll!=0L;
                else rr = rr || x.ll!=0L;
            }
            return new Result( rr );
      }
        throw new BadFormulaException( MSG_INTERNAL_ERROR );
    }

    /* ==============================================================
     * Evaluate
     */
    
    private static char get1() {
        char c;
        while( (c=buf[nBuf++]) == ' ' || c == '\t' || c == '\n' ) ;
        return c;
    }
    
    private static void unget1() { nBuf--; }
    
    private static boolean isLetter( char c ) {
        return
        (c >= 'a' && c <= 'z') ||
        (c >= 'A' && c <= 'Z') ||
        (c >= '_');
    }
    
    public static Result Evaluate( String ss ) throws BadFormulaException {
        Result rr;
        
        // copy src to the char array
        int len = ss.length();
        buf = new char[len+1];
        ss.getChars( 0, len, buf, 0 );
        buf[len] = '\0';
        
        nBuf = 0;
        
        if( get1() == '=' ) {
            rr = f_comp();
            if( get1() != '\0' ) throw new BadFormulaException( "Bad expression" );
        }
        else {
            unget1();
            try {
                rr = f_const();
                if( get1() != '\0' ) throw new BadFormulaException();	// the constant isn't properly ended
            } catch( BadFormulaException e ) {
                rr = new Result( ss );			// will be a string then
            }
        }
        
        buf = null;
        return rr;
    }
    
    private static Result f_comp() throws BadFormulaException {	// must be called 'f_concat()'
        Result a;
        a = f_comp1();
        while( get1() == '&' ) {
            a = new Result( OPER_CONCAT, a, f_comp1() );
        }
        unget1();
        return a;
    }
    
    private static Result f_comp1() throws BadFormulaException {
        Result a,b;
        char c;
        a = f_add();
        while( (c=get1()) == '=' || c == '<' || c == '>' || c == '!' ) {
            byte oper = OPER_EQ;
            if( c == '!' ) {
                if( buf[nBuf++] != '=' )
                    throw new BadFormulaException("Expected '!=' statement");
                oper = OPER_NE;
            }
            else if( c != '=' ) {
                if( buf[nBuf] == '=' ) {
                    nBuf++;
                    oper = c == '>' ? OPER_GE : OPER_LE;
                }
                else
                    oper = c == '>' ? OPER_GT : OPER_LT;
            }
            b = f_add();
            a = new Result( oper, a, b );
        }
        unget1();
        return a;
    }
    
    private static Result f_add() throws BadFormulaException {
        Result a,b;
        char c;
        a = f_mul();
        while ( (c=get1()) == '+' || c == '-' ) {
            b = f_mul();
            a = new Result( c == '+' ? OPER_ADD : OPER_SUB, a, b );
        }
        unget1();
        return a;
    }
    
    private static Result f_mul() throws BadFormulaException {
        Result a,b;
        char c;
        
        a = f_pow();
        while ( (c=get1()) == '*' || c == '/' ) {
            b = f_pow();
            a = new Result( c == '*' ? OPER_MUL : OPER_DIV, a, b );
        }
        unget1();
        return a;
    }
    
    private static Result f_pow() throws BadFormulaException {
        Result a,b;
        char c;
        a = f_uminus();
        while ( get1() == '^' ) {
            b = f_pow();                                                // from right to left
            a = new Result( OPER_POW, a, b );
        }
        unget1();
        return a;
    }
    
    private static Result f_uminus()  throws BadFormulaException {
        if( get1() == '-' )
            return new Result( OPER_UMIN, f_brackets(), null );
        unget1();
        return f_brackets();
    }
    
    private static Result f_brackets()  throws BadFormulaException {
        if( get1() == '(' ) {
            Result a = f_comp();
            if( get1() != ')' ) throw new BadFormulaException("missing ')'");
            return new Result( OPER_BRACK, a, null );
        }
        unget1();
        return f_name();
    }
    
    private static Result f_const() throws BadFormulaException {
        char c;
        
        c = get1();
        unget1();
        
        if( c == '\'' )
            return f_const_string1();
        
        if( c == '-' ) {
            // we shouldn't produce expression tree here (I could parse to tree then calculate it)
            c = get1();
            Result rr = f_const_numeric();
            if( rr.type == TYPE_LONG || rr.type == TYPE_BCD ) rr.ll = -rr.ll;
            return rr;
        }
        
        if( Character.isDigit( c ) ) {
            // check if it's a date
            int nn = nBuf;
            while( (c=buf[nn++]) != '\0' ) {
                if( c == '/' || c == ':' ) {
                    return f_const_date();
                }
            }
            return f_const_numeric();
        }
        
        StringBuffer sb = new StringBuffer(10);
        
        while(
        isLetter( c=buf[nBuf++] ) ||
        Character.isDigit( c ) ) {
            sb.append( c );
        }
        unget1();
        String name = sb.toString().toUpperCase();
        
        if( name.compareTo( "TRUE" ) == 0 )
            return new Result( true );
        if( name.compareTo( "FALSE" ) == 0 )
            return new Result( false );
        
        throw new BadFormulaException();	// not a constant
    }
    
    private static final String MSG_FUNCTION = "missing ')' in function";
    
    private static Result f_name() throws BadFormulaException {
        char c;
        
        c = get1();
        if( c == '#' ) {
            Result rr = f_const_date();
            if( get1() != '#' )
                throw new BadFormulaException("expected trailing '#' in date constant");
            return rr;
        }
        
        unget1();
        
        if( c == '"' )
            return f_const_string();
        
        if( Character.isDigit( c ) )
            return f_const_numeric();
        
        StringBuffer sb = new StringBuffer(10);
        
        // address
        while(
        isLetter( c=buf[nBuf++] ) ||
        Character.isDigit( c ) ||
        c == '$' ||
        c == ':' ) {
            sb.append( c );
        }
        unget1();
        
        if( sb.length() <= 0 )
            throw new BadFormulaException("bad formula - expected a constant");
        
        String name = sb.toString().toUpperCase();
        
        if( get1() == '(' ) {
            // function
            Vector vv = new Vector(3);
            if( (c=get1()) != ')' ) unget1();
            while( c != ')' ) {
                if( c == '\0' ) throw new BadFormulaException(MSG_FUNCTION);
                vv.addElement( f_comp() );
                c = get1();
                if( c != ',' && c != ')' ) throw new BadFormulaException(MSG_FUNCTION);
            }

            int nn = vv.size();
            Result funcargs[] = new Result[nn];
            for( int i=0; i<nn; i++ )
                funcargs[i] = (Result) vv.elementAt(i);
            
            byte functype = (byte) findFunction( name );
            if( functype == -1 )
                throw new BadFormulaException( "Unknown function <" + name + ">" );
            return new Result( functype, funcargs );
        }
        unget1();
        
        if( name.compareTo( "TRUE" ) == 0 )
            return new Result( true );
        if( name.compareTo( "FALSE" ) == 0 )
            return new Result( false );
        
        // cell address
        Result rr = parseAddress( name );
        
        return rr;
        
    }
    
    private static Result f_const_string() throws BadFormulaException {
        char c;
        Result rr;
        StringBuffer sb = new StringBuffer(20);
        nBuf++;					// skip the first '"'
        while( (c=buf[nBuf++]) != '"' ) {
            if( c == '\0' ) {
                unget1();
                break;
            }
            sb.append( c );
        }
        rr = new Result( sb.toString() );
        return rr;
    }
    
    private static Result f_const_string1() throws BadFormulaException {
        char c;
        Result rr;
        StringBuffer sb = new StringBuffer(20);
        nBuf++;					// skip the first '\''
        while( (c=buf[nBuf]) != '\0' ) {
            nBuf++;
            sb.append( c );
        }
        rr = new Result( sb.toString() );
        return rr;
    }
    
    private static Result f_const_numeric() throws BadFormulaException {
        char c;
        StringBuffer sb = new StringBuffer(20);
        while( ((c=buf[nBuf++]) >= '0' && c <= '9') || c == '.' ) {
            sb.append( c );
        }
        String ss = sb.toString();
        if( c == 'L' || c == 'l' )
            return new Result( TYPE_LONG, Long.parseLong( ss ) );
        unget1();
        return createFloat( MathFP.toFP(ss) );
    }
    
    private static final String MSG_DATE = "bad date format (must be DD/MM/YYYY)";
    
    private static Result f_const_date() throws BadFormulaException {
        
        char c;
        StringBuffer sb = new StringBuffer(20);
        while( ((c=buf[nBuf++]) >= '0' && c <= '9') || c == '/' || c == ':' || c == ' ' ) {
            sb.append( c );
        }
        unget1();
        String ss = sb.toString();
        
        int nn1 = ss.indexOf( '/' );
        int nn2 = ss.indexOf( '/', nn1+1 );
        int nn3 = ss.indexOf( ' ', nn2+1 );
        int nn4 = ss.indexOf( ':', nn3+1 );
        int nn5 = ss.indexOf( ':', nn4+1 );
        
        int day = 0;
        int month = 0;
        int year = 0;
        
        int hour = 0;
        int min = 0;
        int sec = 0;
        
        Calendar cd = Calendar.getInstance();
        
        try {
            if( nn1 > 0 && nn2 > 0 ) {
                day = Integer.parseInt( ss.substring( 0,nn1 ) );
                month = Integer.parseInt( ss.substring( nn1+1, nn2 ) );
                year = Integer.parseInt( ss.substring( nn2+1, nn3>0 ? nn3 : ss.length() ) );
                
                if( day < 0 || day > 31 || month <= 0 || month > 12 || year < 1970 )
                    throw new BadFormulaException(MSG_DATE);
                
                cd.set( Calendar.DAY_OF_MONTH, day );
                cd.set( Calendar.MONTH, month-1 );
                cd.set( Calendar.YEAR, year );
            } else {
                cd.setTime( new Date() );
            }
            if( nn4 > 0 ) {
                hour = Integer.parseInt( ss.substring( nn3+1,nn4 ) );
                min = Integer.parseInt( ss.substring( nn4+1, nn5>0 ? nn5 : ss.length() ) );
            }
            if( nn5 > 0 ) {
                sec = Integer.parseInt( ss.substring( nn5+1 ) );
            }
            cd.set( Calendar.HOUR_OF_DAY, hour );
            cd.set( Calendar.MINUTE, min );
            cd.set( Calendar.SECOND, sec );
        }
        catch( Exception e ) {
            throw new BadFormulaException(MSG_DATE);
        }
        
        return new Result( TYPE_DATE, cd.getTime().getTime() );
    }
    
    public static boolean intersect(
    short i1, short j1, short i2, short j2,
    short _i1, short _j1, short _i2, short _j2
    ) {
        short x1 = j2 < _j2 ? j2 : _j2;
        short y1 = i2 < _i2 ? i2 : _i2;
        short x2 = j1 > _j1 ? j1 : _j1;
        short y2 = i1 > _i1 ? i1 : _i1;
        return x1 <= x2 && y1 <= y2;
    }
    
    // returns true if changed
    public boolean shiftReferences( short I1, short J1, short dI, short dJ ) {
        
        boolean rr=false,rr1=false,rr2=false,inter=false;
        
        for( int i=0; i<funcargs.length; i++ ) {
	    rr1 = funcargs[i].shiftReferences( I1,J1, dI,dJ );
            rr = rr || rr1;
	}
        
        if( type != TYPE_RANGE ) return rr || rr1 || rr2;
        
        //System.out.println("range encountered");
        
        short _i1=i1,_j1=j1,_i2=i2,_j2=j2;	// old values
        
        if( dI < 0 ) {
            
            if( i1 >= I1 && i2 < I1-dI ) {	// range within deleted area
                type = TYPE_ERROR;
                return true;
            }
            
            inter = intersect( i1,j1,i2,j2, I1,J1,(short)(I1-dI),j2 );	// check _before_ deleting
            
            if( i1 >= I1-dI ) i1 += dI;
            else if( i1 > I1 ) i1 = I1;
            
            if( i2 >= I1-dI ) i2 += dI;
            else if( i2 > I1 ) i2 = I1;
        }
        else if( dI > 0 ) {
            if( i1 >= I1 ) i1 += dI;
            if( i2 >= I1 ) i2 += dI;
            inter = intersect( i1,j1,i2,j2, I1,J1,(short)(I1+dI),j2 );	// check _after_ inserting
        }	// dI = 0
        
        if( dJ < 0 ) {
            
            if( j1 >= J1 && j2 < J1-dJ ) {
                type = TYPE_ERROR;
                return true;
            }
            
            inter = intersect( i1,j1,i2,j2, I1,J1,i2,(short)(J1-dJ) );	// check _before_ deleting
            if( j1 >= J1-dJ ) j1 += dJ;
            else if( j1 > J1 ) j1 = J1;
            
            if( j2 >= J1-dJ ) j2 += dJ;
            else if( j2 > J1 ) j2 = J1;
        }
        else if( dJ > 0 ) {
            if( j1 >= J1 ) j1 += dJ;
            if( j2 >= J1 ) j2 += dJ;
            inter = intersect( i1,j1,i2,j2, I1,J1,i2,(short)(J1+dJ) );
        }	// dJ = 0
        
        //System.out.println("_i1="+_i1+" _j1="+_j1+" _i2="+_i2+" _j2="+_j2);
        rr = inter || i1 != _i1 || j1 != _j1 || i2 != _i2 || j2 != _j2;	// formula has changed ?
        //System.out.println("i1="+i1+" j1="+j1+" i2="+i2+" j2="+j2);
        return rr;
    }

    // shifts references according to absolute flags (being used in copy/paste operations)
    public boolean moveReferences( int di, int dj ) {
        boolean rr = false, rr1;
        for( int i=0; i<funcargs.length; i++ ) {
            rr1 = funcargs[i].moveReferences( di, dj );
            rr = rr || rr1;                 // !!! rr1 is used intentionally, don't optimize, we _need_ to call mobeReferences again even rr already is true
        }                                   // I already fixed this 2 times :))
        if( type == TYPE_RANGE ) {
            if( (absolute&ABS_I1) == 0 ) i1 += di;
            if( (absolute&ABS_J1) == 0 ) j1 += dj;
            if( (absolute&ABS_I2) == 0 ) i2 += di;
            if( (absolute&ABS_J2) == 0 ) j2 += dj;
            rr = (absolute&(ABS_I1|ABS_J1|ABS_I2|ABS_J2)) != (ABS_I1|ABS_J1|ABS_I2|ABS_J2);   // only of all abs there won't be changes in references
        }
        return rr;
    }

    private static int findFunction( String ss ) {
        ss = ss.toUpperCase();
        for( int i=0; i<FUNCTION_NAMES.length; i++ )
            if( ss.compareTo( FUNCTION_NAMES[i] ) == 0 )
                return i;
        // not found
        return -1;
    }
    
    public String getFunctionName() {
        return FUNCTION_NAMES[(int)ll];
    }

    // to do - combine it with Result(InputStrem) to prevent double code (absolute)
    public static Result createCell( int i, int j, String formula, Result rr ) {
        Result cell = new Result( TYPE_CELL );
        cell.i1 = (short)i; cell.j1 = (short)j;
        cell.funcargs = new Result[1];
        cell.funcargs[0] = rr;
        cell.str = formula;
        if( formula.length() > 0 && formula.charAt(0) == '=' )
            cell.absolute |= FORMULA;
        return cell;
    }
    
    public static Result createFloat( long ll ) {
        Result rr = new Result( TYPE_BCD, ll );
        return rr;
    }

    public boolean isFormula() {
        return (absolute & FORMULA) != 0;
    }
    
    // false for empty cell without formula
    public boolean hasFormula() {
        return str != EMPTY_FORMULA;
    }
    
    public boolean isEmptyCell() {
        return funcargs[0].type == TYPE_EMPTY;      // don't compare with CELL_EMPTY since there may be an expression
    }

    public int hashCode() {
        return (i1<<16)|j1;
    }
    
    public boolean equals( Object obj ) {
        return hashCode() == obj.hashCode();
    }
}
