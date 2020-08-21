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

import javax.microedition.midlet.*;
import javax.microedition.lcdui.*;
import javax.microedition.rms.*;
import java.io.*;

/** the canvas class to perform all the user iteractions
 */
public final class MicroCalcCanvas extends Canvas implements CommandListener {
    
    /* ======================================================
     * menu
     */
    static final byte CELL_EDIT		= 1;
    static final byte CELL_CLEAR	= 2;
    static final byte CELL_COPY		= 3;
    static final byte CELL_PASTE	= 4;
    
    static final byte COLUMN_WIDTH	= 5;
    static final byte SHEET_NEW		= 6;
    static final byte SHEET_SAVE	= 7;
    static final byte SHEET_LOAD	= 8;
    static final byte SHEET_DELETE	= 9;
    
    static final byte ABOUT		= 13;
    static final byte SHEET_NEW1	= 19;
    
    static final byte COLUMN_INSERT	= 20;
    static final byte COLUMN_DELETE	= 22;
    
    static final byte ROW_INSERT	= 24;
    static final byte ROW_DELETE	= 26;
    
    static final byte SHEET_RESIZE	= 28;
    
    static final byte FORMAT_NUMBER	= 30;
    static final byte FORMAT_DATETIME	= 31;
    static final byte FORMAT_ALIGN	= 32;
    static final byte FORMAT_CLEAR	= 33;
    static final byte FORMAT_FONT	= 34;
    
    static final byte CONFIRM_OK	= 50;
    static final byte CONFIRM_CANCEL	= 51;
    static final byte COMMAND_EXIT	= 52;
    
    static final byte EXIT              = 100;
    
    private static final String startMenuLabels[] = {
        "Start menu",
        "Load sheet",
        "New sheet",
        "About",
        "Exit"
    };
    private static final short startMenuCodes[] = {
        0,
        SHEET_LOAD | Menu.MASK_PROMPT,
        SHEET_NEW1 | Menu.MASK_PROMPT,
        ABOUT | Menu.MASK_PROMPT,
        EXIT | Menu.MASK_PROMPT
    };
    private static final byte startMenuParents[] = {
        -1,
        0,
        0,
        0,
        0
    };
    
    private static final String confirmMenuLabels[] = {
        "Sheet changed!",
        "Discard changes",
        "Cancel"
    };
    private static final short confirmMenuCodes[] = {
        0,
        CONFIRM_OK,
        CONFIRM_CANCEL
    };
    private static final byte confirmMenuParents[] = {
        -1,
        0,
        0
    };
    private byte confirmCode;		// code being confirmed
    
    private static final String defaultLabels[] = {
        "Top menu",
        
        "Cell...",		// main
        "Column...",
        "Row...",
        "Sheet...",
        "About",
        "Exit",
        
        "Edit",			// cell
        "Clear",
        "Copy",
        "Paste",
        "Format...",
        
        "Width",		// column
        "Insert",
        "Delete",
        
        "Insert",		// row
        "Delete",
        
        "New",			// sheet
        "Save",
        "Load",
        "Delete",
        "Resize",
//        "-Test-",
        
//        "Align",		// cell...format
//        "Font",
        "Number",
        "Date/Time",
        "Clear"
        
    };
    private static final byte defaultParents[] = {
        -1,
        
        0,	// top menu
        0,
        0,
        0,
        0,
        0,
        
        1,	// cell menu
        1,
        1,
        1,
        1,
        
        2,	// column
        2,
        2,
        
        3,	// row
        3,
        
        4,	// sheet
        4,
        4,
        4,
        4,
//        4,
        
//        10,	// cell.../format
//        10,
        11,
        11,
        11
    };
    
    private static final short defaultCodes[] = {
        0,
        
        0,	// top menu
        0,
        0,
        0,
        ABOUT | Menu.MASK_PROMPT,
        EXIT | Menu.MASK_PROMPT,
        
        CELL_EDIT | Menu.MASK_PROMPT,		// cell menu
        CELL_CLEAR,
        CELL_COPY,
        CELL_PASTE,
        0,
        
        COLUMN_WIDTH | Menu.MASK_PROMPT,	// column
        COLUMN_INSERT | Menu.MASK_PROMPT,
        COLUMN_DELETE | Menu.MASK_PROMPT,
        
        ROW_INSERT | Menu.MASK_PROMPT,
        ROW_DELETE | Menu.MASK_PROMPT,
        
        SHEET_NEW | Menu.MASK_PROMPT,		// sheet menu
        SHEET_SAVE | Menu.MASK_PROMPT,
        SHEET_LOAD | Menu.MASK_PROMPT,
        SHEET_DELETE | Menu.MASK_CONFIRM,
        SHEET_RESIZE | Menu.MASK_PROMPT,
//        TEST | Menu.MASK_PROMPT,
        
//        FORMAT_ALIGN | Menu.MASK_PROMPT,	// cell.../format
//        FORMAT_FONT | Menu.MASK_PROMPT,
        FORMAT_NUMBER | Menu.MASK_PROMPT,
        FORMAT_DATETIME | Menu.MASK_PROMPT,
        FORMAT_CLEAR
    };
    
    private static final Menu defaultMenu = new Menu( defaultLabels, defaultCodes, defaultParents );
    private static final Menu confirmMenu = new Menu( confirmMenuLabels, confirmMenuCodes, confirmMenuParents );
    public static final Menu startMenu = new Menu( startMenuLabels, startMenuCodes, startMenuParents );
    
    // =============== end of menu
    
    // directions of cursor's movement
    public static final byte CURSOR_UP		= 0;
    public static final byte CURSOR_DOWN	= 1;
    public static final byte CURSOR_LEFT	= 2;
    public static final byte CURSOR_RIGHT	= 4;
    public static final byte CURSOR_NONE	= 5;
    
    private static final String DEFAULT_WIDTH  = "######";
    private static final String DEFAULT_HEADER = "##";
    
    // commands
//    private static final Command cmdExit	= new Command("Exit", Command.EXIT, 2);
//    private static final Command cmdAbout	= new Command("Help", Command.HELP, 3);
    private static final Command cmdMenu	= new Command("Menu", Command.OK, 1);
    
    // Commands used in the forms
    private static final Command cancelEditCommand 	= new Command("Cancel", Command.BACK, 2);
    private static final Command okEditCommand 		= new Command("OK", Command.OK, 1);
    private static final Command cmdDelete 		= new Command("Delete", Command.SCREEN, 3);
    
/*
    public static final int DEFAULT_ROWS	= 32;
    public static final int DEFAULT_COLUMNS	= 8;
 */

    private int sizeX, sizeY;		// screen size
    
    // displayable window in the sheet coordinates
    private int windowX1, windowY1, windowX2, windowY2;
    
    // cursor's location, I = Y, J = X
    int cursorI, cursorJ;
    // cursor direction
    int dirI, dirJ;
    private int clipboardI=-1,clipboardJ=-1;
    
    // font attribs
    Font font;
    int fontHeight;
    int defaultColumnWidth;
    int headerWidth;
    
    Display display;
    
    MicroCalc parent;
    private Alert directions;		// for about screen
    
    private Form frmNewSheet, frmResize;
    private Form frmFormatNumber, frmFormatAlign, frmFormatDate;
    private Form frmInsertColumn, frmInsertRow;
    private Form frmDeleteColumn, frmDeleteRow;
    private TextBox edit;	// formula editor
    public TextBox editName;	// to prompt sheet name
    private TextBox editColumnWidth;	// to prompt column width
    private TextBox editRowHeight;	// to prompt row height
    public List selectList;	// to select the sheet to load
    
    public Sheet sheet;                      // current sheet (only one loaded in memmory)
    private String statusLine;               // status line to be displayed if not null
    private byte statusLinePart;             // to scroll status line horizontally
    // following booleans are candidates to bit mask status
    private boolean statusGold;              // true if the GOLD (#) key has been is pressed and we are in gold mode
    private boolean statusNumeric;           // true when we are in NUMERIC mode when a decimal constant is being entered into the current cell
    private String numericLine;              // keeps user input line in NUMERIC mode
    private boolean statusInfo;              // true when cell info is displayed in the status line
    private boolean statusMem;               // true when mem info is displayed in the status line
    
    public MicroCalcCanvas( MicroCalc _parent ) {
        
        parent = _parent;
        
        display = Display.getDisplay( parent );
        
        font = Font.getFont( Font.FACE_SYSTEM, Font.STYLE_PLAIN, Font.SIZE_SMALL );
        fontHeight = font.getHeight();
        defaultColumnWidth = font.stringWidth( DEFAULT_WIDTH );
        headerWidth = font.stringWidth( DEFAULT_HEADER );
        
        sizeX = getWidth();
        sizeY = getHeight();
        
        sheet = new Sheet( Sheet.newRows, Sheet.newColumns, defaultColumnWidth+3, fontHeight+3 );
        
        setWindow( 0, 0, CURSOR_DOWN, CURSOR_RIGHT );
        
//        addCommand( cmdExit );
        addCommand( cmdMenu );
        
        setCommandListener( this );
        
    }
    
    public void setWindow( int _cursorI, int _cursorJ, int _dirI, int _dirJ ) {
        
        int sizex = sizeX - headerWidth;
        int sizey = sizeY - fontHeight;
        
        cursorI = _cursorI; cursorJ = _cursorJ;
        dirI = _dirI; dirJ = _dirJ;
        
        if( dirI == CURSOR_DOWN ) {
            windowY1 = sheet.cellY( cursorI );
            windowY2 = windowY1 + sizey;
        } else if( dirI == CURSOR_UP ) {
            windowY2 = sheet.cellY( cursorI + 1 );
            windowY1 = windowY2 - sizey;
        }
        
        if( dirJ == CURSOR_RIGHT ) {
            windowX1 = sheet.cellX( cursorJ );
            windowX2 = windowX1 + sizex;
        } else if( dirJ == CURSOR_LEFT ) {
            windowX2 = sheet.cellX( cursorJ + 1 );
            windowX1 = windowX2 - sizex;
        }
        
        if( windowX1 < 0 ) { windowX1 = 0; windowX2 = sizex; }
        if( windowY1 < 0 ) { windowY1 = 0; windowY2 = sizey; }
        
        int nn;
        if( windowX2 > (nn=sheet.cellX(sheet.columns)) && nn > sizex ) { windowX2 = nn; windowX1 = windowX2 - sizex; }
        if( windowY2 > (nn=sheet.cellY(sheet.rows)) && nn > sizey ) { windowY2 = nn; windowY1 = windowY2 - sizey; }
        
    }
    
    private void repaintCell( int i, int j ) {
        repaint( headerWidth+sheet.cellX(j)-windowX1+1, fontHeight+sheet.cellY(i)-windowY1+1, sheet.columnWidth[j], sheet.rowHeight[i] );
    }
    
    // returns true if total repaint() was called
    private boolean setCursor( int I, int J ) {
        int diri = CURSOR_NONE, dirj = CURSOR_NONE;
        int oldi = cursorI, oldj = cursorJ;
        boolean willRepaint = false;
        
        if( I < 0 || I >= sheet.rows || J < 0 || J >= sheet.columns )
            return false;
        
        if( sheet.cellY( I ) < windowY1 ) {
            willRepaint = true;
            diri = CURSOR_UP;
        } else if( sheet.cellY( I+1 ) >= windowY2 ) {
            willRepaint = true;
            diri = CURSOR_DOWN;
        }
        
        if( sheet.cellX( J ) < windowX1 ) {
            willRepaint = true;
            dirj = CURSOR_LEFT;
        } else if( sheet.cellX( J+1 ) >= windowX2 ) {
            willRepaint = true;
            dirj = CURSOR_RIGHT;
        }
        
        if( willRepaint ) {
            setWindow( I, J, diri, dirj );
            repaint();
        }
        else if( oldi != I || oldj != J ) {
            repaintCell( oldi, oldj );
            cursorI = I; cursorJ = J;
            repaintCell( cursorI, cursorJ );
        }
        return willRepaint;
    }
    
    /*
     * Handle a key pressed event.
     */
    protected void keyPressed(int keyCode) {
        
//        Runtime.getRuntime().gc();
        
        boolean willRepaint = false;
        
        if( statusNumeric ) {
            switch( keyCode ) {
                case KEY_NUM0:
                case KEY_NUM1:
                case KEY_NUM2:
                case KEY_NUM3:
                case KEY_NUM4:
                case KEY_NUM5:
                case KEY_NUM6:
                case KEY_NUM7:
                case KEY_NUM8:
                case KEY_NUM9:
                    numericLine += (char) ('0' + (keyCode - KEY_NUM0));
                    break;
                case KEY_STAR:
                    if( numericLine.indexOf( '.' ) >= 0 )
                        numericLine = numericLine.charAt(0) == '-' ? "" : "-";
                    else
                        numericLine += '.';
                    break;
                case KEY_POUND:         // end of NUMERIC mode
                    setStatusLine( null );      // to clear and repaint status line
                    if( numericLine.length() != 0 ) {           // don't clear cell if input is empty
                        try {
                            sheet.setFormula( cursorI, cursorJ, numericLine );
                            sheet.calculateDepended( cursorI, cursorJ, true );
                            fitCell( cursorI, cursorJ );
                            sheet.changed = true;
                            repaint();                      // TODO: only changed cells should be repainted (and probably fitted)
                        } 
                        catch( Exception e ) { }
                    }
//                    repaintCell( cursorI, cursorJ );
                    statusNumeric = false;
                    numericLine = null;
                    return;
            }
            setStatusLine( "NUM: " + numericLine );    // to repaint status line
            return;
        }
        
        // GOLD mode menu
        if( statusGold ) {
            int old;                            // used in PgUp/PgDn
            statusGold = false;                 // quiting gold mode by pressing any key
            switch( keyCode ) {
                case KEY_NUM0:                 // enter NUMERIC mode
                    statusNumeric = true;
                    statusInfo = false;
                    numericLine = "";
//                    statusline = "NUM: ";
                    break;
                case KEY_NUM1:                 // copy cell
                    clipboardI = cursorI;
                    clipboardJ = cursorJ;
                    break;
                case KEY_NUM2:                 // paste cell
                    if( clipboardI >= 0 && clipboardJ >= 0 ) {
                        sheet.copyCell1( clipboardI, clipboardJ, cursorI-clipboardI, cursorJ-clipboardJ );
                        willRepaint = true;
                    }
                    break;
// comment this for LITE (optional)
                case KEY_NUM3:                      // page up
                    old = windowY1;
                    while( cursorI > 0 && old == windowY1 )
                        setCursor( cursorI-1, cursorJ );
                    willRepaint = true;
                    break;
                case KEY_NUM6:                      // page up
                    old = windowY1;
                    while( cursorI < sheet.rows-1 && old == windowY1 )
                        setCursor( cursorI+1, cursorJ );
                    willRepaint = true;
                    break;
                case KEY_NUM4:                      // page left
                    old = windowX1;
                    while( cursorJ > 0 && old == windowX1 )
                        setCursor( cursorI, cursorJ-1 );
                    willRepaint = true;
                    break;
                case KEY_NUM5:                      // page right
                    old = windowX1;
                    while( cursorJ < sheet.columns-1 && old == windowX1 )
                        setCursor( cursorI, cursorJ+1 );
                    willRepaint = true;
                    break;

                case KEY_NUM8:                      // toggle on/off mem monitor
                    statusMem = !statusMem;
                    break;
                case KEY_NUM9:                      // toggle on/off status info
                    statusInfo = !statusInfo;
                    break;
            }
            String statusline = null;
            if( statusInfo )
                    statusline = getStatusInfo( cursorI, cursorJ );
            else if( statusNumeric )
                    statusline = "NUM: ";
            else if( statusMem )
                    statusline = getMemoryInfo();
            setStatusLine( statusline );
            if( willRepaint ) repaint();                    // don't forget about repaint order - ie Nokia 7650 ignores first repaint request, only last one performed
            return;
        }

        // enter GOLD mode from normal
        if( keyCode == KEY_POUND ) {
            statusGold = true;
            setStatusLine( "GOLD? " );
            return;
        }
        
        // clear the status line and mode
        statusGold = statusNumeric = false;

        int action = getGameAction(keyCode);
        
        int ii = cursorI, jj = cursorJ;
        
        switch (action) {
            
            case DOWN:
                ii++;
                break;
                
            case UP:
                ii--;
                break;
                
            case RIGHT:
                jj++;
                break;
                
            case LEFT:
                jj--;
                break;
                
            case FIRE:
                editCell( cursorI, cursorJ );
                return;                             // no repaint needed, switching screen to edit box
                
            default:
                return;                             // unknown keycode, do nothing
        }
        
        // only cursor keys may reach here
        willRepaint = setCursor( ii, jj );
        
        // clear the status line or display formula if cell info mode
        if( statusInfo || statusMem ) {
            statusLine = statusInfo ? getStatusInfo( cursorI, cursorJ ) : getMemoryInfo();
            if( !willRepaint )
                setStatusLine( statusLine );        // force to repaint
//            setStatusLine( getStatusInfo( cursorI, cursorJ ), !willRepaint );
        }
/*        
        if( willRepaint ) 
            repaint();        // ugly - again, may be 3rd time (7650 draws only last repaint)
 */
    }
    
    protected void keyRepeated( int keyCode ) {
        keyPressed( keyCode );
    }

// comment this for LITE version
    protected void pointerPressed( int x, int y ) {
        
//        Runtime.getRuntime().gc();
        
        if( x >= sizeX || y >= sizeY )
            return;
        
        x += windowX1-headerWidth; y += windowY1-fontHeight;
        if( x < 0 || y < 0 ) return;
        
        int I=sheet.rows-1, J=sheet.columns-1;
        
        int Y = 0;
        for( int i=0; i<sheet.rows; i++ ) {
            Y += sheet.rowHeight[i];
            if( y < Y ) {
                I = i;
                break;
            }
        }
        int X = 0;
        for( int j=0; j<sheet.columns; j++ ) {
            X += sheet.columnWidth[j];
            if( x < X ) {
                J = j;
                break;
            }
        }
        if( cursorI == I && cursorJ == J ) {
            editCell( cursorI, cursorJ );
        }
        else {
            setCursor( I, J );
        }
    }
    
    protected void paint( Graphics g ) {
        
        int clipx = g.getClipX(), clipy = g.getClipY();
        int clipdx = g.getClipWidth(), clipdy = g.getClipHeight();
        
//        int wX1 = windowX1 + (clipx - headerWidth);
//        int wY1 = windowY1 + (clipy - fontHeight);
        
        // flag to draw row names
        boolean rowHeader = intersect( 
            clipx,clipy, clipx+clipdx,clipy+clipdy,
            headerWidth,0, sizeX,fontHeight );
        // flag to draw column names
        boolean columnHeader = intersect( 
            clipx,clipy, clipx+clipdx,clipy+clipdy,
            0,fontHeight, headerWidth,sizeY );
        boolean drawCells = intersect(
            clipx,clipy, clipx+clipdx,clipy+clipdy,
            headerWidth+1,fontHeight+1, sizeX,sizeY );

        boolean drawStatusLine = (statusLine != null) && rowHeader;      // status line will be also painted
                                                                            // if true then columnHeader is also true
        // redraw area in abs coords 
        int wX1 = windowX1 + (clipx - headerWidth);         
        int wY1 = windowY1 + (clipy - fontHeight);
        int wX2 = wX1 + clipdx;
        int wY2 = wY1 + clipdy;

        if( rowHeader ) wX1 = windowX1;
        if( columnHeader ) wY1 = windowY1;

        int i1=-1, i2=sheet.rows-1;
        int j1=-1, j2=sheet.columns-1;
        int nn;

        /* find i1,i2, j1,j2 - visible cells area
         */
        int X = 0;
        for( int j=0; j<sheet.columns; j++ ) {
            if( X >= wX2 ) {
                j2 = j-1;
                break;
            }
            nn = sheet.columnWidth[j];
            if( j1 == -1 && (X >= wX1 || X+nn > wX1) )
                j1 = j;
            X += nn;
        }
        
        int Y = 0;
        for( int i=0; i<sheet.rows; i++ ) {
            if( Y >= wY2 ) {
                i2 = i-1;
                break;
            }
            nn = sheet.rowHeight[i];
            if( i1 == -1 && (Y >= wY1 || Y+nn > wY1) )
                i1 = i;
            Y += nn;
        }
        
        g.setColor( 0xffffff );
        g.fillRect( 0,0, sizeX, sizeY );
        
        g.setFont( font );
        
        // header
        if( rowHeader ) {
            
/*
            g.setColor( 0xffffff );
            g.fillRect( 0,0, sizeX,fontHeight );
 */
            g.setColor( 0x000000 );
            g.drawLine( headerWidth,fontHeight, sizeX-1,fontHeight );
            
            // paint the status line
            if( drawStatusLine ) {
                g.drawString( statusLine, 2, fontHeight, Graphics.LEFT|Graphics.BOTTOM );
            }
            // paint column names
            else {
                g.setClip( headerWidth,0, sizeX-headerWidth,fontHeight );
                X = sheet.cellX( j1 ) - windowX1 + headerWidth;
                int X1;
                g.drawLine( X,0, X,fontHeight-1 );
                for( int j=j1; j<=j2; j++ ) {
                    X1 = X + sheet.columnWidth[j];
                    g.drawLine( X1,0, X1,fontHeight-1 );
                    g.drawString( Result.columnName( j ), (X+X1)/2, fontHeight, Graphics.BOTTOM|Graphics.HCENTER );
                    X = X1;
                }
            }
            
            g.setClip( 0,0, sizeX,sizeY );
            
        }
        
        // header
        if( columnHeader ) {
            
/*
            g.setColor( 0xffffff );
            g.fillRect( 0,fontHeight+1, headerWidth,sizeY );
 */
            
            g.setColor( 0x000000 );
            g.drawLine( headerWidth, fontHeight, headerWidth,sizeY-1 );
            
            g.setClip( 0,fontHeight, headerWidth,sizeY-fontHeight );
            
            Y = sheet.cellY( i1 ) - windowY1 + fontHeight;
            int Y1;
            g.drawLine( 0,Y, headerWidth-1,Y );
            for( int i=i1; i<=i2; i++ ) {
                Y1 = Y + sheet.rowHeight[i];
                g.drawLine( 0,Y1, headerWidth-1,Y1 );
                g.drawString( String.valueOf(i+1), headerWidth-1, Y+2, Graphics.TOP|Graphics.RIGHT );
                Y = Y1;
            }
        }
        
        if( drawCells ) {
            for( int i=i1; i<=i2; i++ ) {
                for( int j=j2; j>=j1; j-- ) {
                    // to intersect with in clipRect()
                    g.setClip( headerWidth+1,fontHeight+1, sizeX-headerWidth-1,sizeY-fontHeight-1 );
                    paintCell( g, i,j, i==cursorI && j==cursorJ );
                }
            }
        }
    }
    
    void paintCell( Graphics g, int I, int J, boolean selected ) {
        int x = sheet.cellX(J) - windowX1 + headerWidth;
        int y = sheet.cellY(I) - windowY1 + fontHeight;
        
        int nx = sheet.columnWidth[J];
        int ny = sheet.rowHeight[I];
        
        int xx = x + nx;
        int yy = y + ny;
        
        boolean rightLine = true;
        
        g.clipRect( x, y+1, nx+1,ny ); 	// to include the left border that may be cleared by the text to the left
        
        //	g.setFont( font );
        
        if( selected ) {
            g.setColor( 0x000000 );
            g.fillRect( x+1,y+1, nx+1,ny-1 ); // to clear right border that may be used by the long text
        }
        //	g.setColor( selected ? 0x000000 : 0xffffff );
        //	g.fillRect( x+1,y+1, nx+1,ny-1 ); // to clear right border that may be used by the long text
        
        Result cell = sheet.getCell( I,J );
        Result value = cell.funcargs[0];
        int formats = (int) cell.ll;                          // format is stored here

        if( !cell.isEmptyCell() ) {
            
            g.setColor( selected ? 0xffffff : 0x000000 );
            String ss = value.toString(false,formats);
            
            // draw the string with proper aligment
            if( value.type == Result.TYPE_STRING ) {
                g.drawString( ss, x+2, y+2, Graphics.TOP|Graphics.LEFT );
                rightLine = !sheet.isEmpty( I,J+1 ) || ( font.stringWidth( ss )+3 <= nx );
            }
            else {
                if( font.stringWidth( ss )+3 > nx ) {
                    ss = DEFAULT_WIDTH;
                }
                g.drawString( ss, xx-1, y+2, Graphics.TOP|Graphics.RIGHT );
            }
        }
        // empty
        else {
            // check if there's text from the left to paint within this cell
            int width = 0;
            for( int j=J-1; j>=0; j-- ) {
                width += sheet.columnWidth[j];
                value = sheet.getCellValue( I,j );
                if( value.type != Result.TYPE_EMPTY ) {
                    if( value.type == Result.TYPE_STRING ) {
                        int ww;
                        String ss = value.toString();
                        if( (ww = font.stringWidth( ss )+3) > width ) {
                            g.setColor( !selected ? 0xffffff : 0x000000 );
                            //	    			g.drawLine( x,y+1, x, yy-1 );	// left line
                            g.setColor( selected ? 0xffffff : 0x000000 );
                            g.drawString( ss, x+2-width, y+2, Graphics.TOP|Graphics.LEFT );
                            rightLine = !sheet.isEmpty( I, J+1 ) || ( ww <= width + nx );
                        }
                    }
                    break;          // break for any non-empty cell
                }
            }
        }
        
        g.setColor( 0x000000 );
        if( rightLine )
            g.drawLine( xx,y+1, xx, yy-1 );	// right line
        g.drawLine( x+1,yy, xx, yy );		// bottom line
        
    }
    
    public void commandAction( Command c, Displayable d ) {

//        Runtime.getRuntime().gc();
        
        try {
            
            Result currentCell = sheet.getCell( cursorI, cursorJ );
            int currentFormat = (int)currentCell.ll;
            
            if( c == cmdMenu ) {
                defaultMenu.start( parent, this );
                return;
            }
            if( d == selectList ) {
                if( c == cmdDelete ) {
                    int nn = selectList.getSelectedIndex();
                    if( nn >= 0 ) {
                        String ss = selectList.getString( nn );
                        try {
                            deleteSheet( ss );
                            selectSheet();
                        }
                        catch( Exception e ) {
                            Alert err = new Alert("Cannot load the sheet, error: ", e.getMessage(), null, AlertType.ERROR);
                            display.setCurrent(err, selectList);
                        }
                    }
                    return;
                }
                if( c == List.SELECT_COMMAND || c == okEditCommand  ) {

                    int nn = selectList.getSelectedIndex();
                    if( nn >= 0 ) {
			parent.loadSheet( selectList.getString(nn) );
                    }
                }
                else
                    setCurrent();
                return;
            }
            if( d == editName ) {
                if( c == okEditCommand ) {
                    sheet.name = editName.getString();
                    sheet.changed = true;
		    parent.saveSheet( null );
                }
                else
                    setCurrent();
                editName = null;
                return;
            }
            if( d == frmNewSheet ) {
                if( c == okEditCommand ) {
                    String sheetName = ((TextField)frmNewSheet.get(0)).getString();
                    int nrows = Integer.parseInt( ((TextField)frmNewSheet.get(1)).getString() );
                    int ncols = Integer.parseInt( ((TextField)frmNewSheet.get(2)).getString() );
                    if( nrows < 2 || ncols < 2 )
                        throw new IOException("Wrong sheet size");
                    saveSheet( sheet );
                    sheet.newRows = nrows;
                    sheet.newColumns = ncols;
                    sheet = new Sheet( nrows, ncols, defaultColumnWidth+3, fontHeight+3 );
                    sheet.name = sheetName;
                }
                frmNewSheet = null;
                setCurrent();
                return;
            }
            if( d == frmInsertRow ) {
                if( c == okEditCommand ) {
                    int nrows = Integer.parseInt( ((TextField)frmInsertRow.get(0)).getString() );
                    boolean resize = ((ChoiceGroup)frmInsertRow.get(1)).getSelectedIndex() == 1;
                    int copyValues = ((ChoiceGroup)frmInsertRow.get(2)).getSelectedIndex();
                    if( nrows < 1 )
                        throw new IOException("Wrong number of rows");
                    sheet.insertCells( cursorI,cursorJ, nrows,0, resize );
                    if( copyValues == 1 && cursorI != 0 ) {
                        for( int i=0; i<nrows; i++ )
                            sheet.copyRow1( cursorI-1, i+1 );
                    }
                    if( copyValues == 2 ) {
                        for( int i=0; i<nrows; i++ )
                            sheet.copyRow1( cursorI+nrows, -i-1 );
                    }
                }
                frmInsertRow = null;
                setCurrent();
                return;
            }
            if( d == frmDeleteRow ) {
                if( c == okEditCommand ) {
                    int nrows = Integer.parseInt( ((TextField)frmDeleteRow.get(0)).getString() );
                    boolean resize = ((ChoiceGroup)frmDeleteRow.get(1)).getSelectedIndex() == 1;
                    if( nrows < 1 )
                        throw new IOException("Wrong number of rows");
                    sheet.insertCells( cursorI,cursorJ, -nrows,0, resize );
                }
                frmDeleteRow = null;
                setCurrent();
                return;
            }
            if( d == frmInsertColumn ) {
                if( c == okEditCommand ) {
                    int ncols = Integer.parseInt( ((TextField)frmInsertColumn.get(0)).getString() );
                    boolean resize = ((ChoiceGroup)frmInsertColumn.get(1)).getSelectedIndex() == 1;
                    if( ncols < 1 )
                        throw new IOException("Wrong number of columns");
                    sheet.insertCells( cursorI,cursorJ, 0, ncols, resize );
                }
                frmInsertColumn = null;
                setCurrent();
                return;
            }
            if( d == frmDeleteColumn ) {
                if( c == okEditCommand ) {
                    int ncols = Integer.parseInt( ((TextField)frmDeleteColumn.get(0)).getString() );
                    boolean resize = ((ChoiceGroup)frmDeleteColumn.get(1)).getSelectedIndex() == 1;
                    if( ncols < 1 )
                        throw new IOException("Wrong number of columns");
                    sheet.insertCells( cursorI,cursorJ, 0, -ncols, resize );
                }
                frmDeleteColumn = null;
                setCurrent();
                return;
            }
            if( d == editColumnWidth ) {
                if( c == okEditCommand ) {
                    int nn = Integer.parseInt( editColumnWidth.getString() );
                    if( nn > 4 ) sheet.columnWidth[cursorJ] = (short) nn;
                }
                editColumnWidth = null;
                setCurrent();
                return;
            }
            if( d == frmResize ) {
                if( c == okEditCommand ) {
                    int nrows = Integer.parseInt( ((TextField)frmResize.get(0)).getString() );
                    int ncols = Integer.parseInt( ((TextField)frmResize.get(1)).getString() );
                    if( nrows < 2 || ncols < 2 )
                        throw new IOException("Wrong sheet size");
                    sheet.resize( nrows, ncols );
                }
                frmResize = null;
                setCurrent();
                return;
            }
            if( d == frmFormatNumber ) {
                if( c == okEditCommand ) {
                    currentFormat |= Result.FORMAT_NUMBER;
                    String ss = ((TextField)frmFormatNumber.get(0)).getString();
                    if( ss.length() != 0 )
                        currentFormat |= (Short.parseShort(ss) << 4);
                    if( ((ChoiceGroup)frmFormatNumber.get(1)).getSelectedIndex() == 0 )
                        currentFormat |= Result.FORMAT_NUMBER_SEP;
                    currentCell.ll = currentFormat;
                    fitCell( cursorI, cursorJ );
                    sheet.changed = true;
                }
                frmFormatNumber = null;
                setCurrent();
                return;
            }
            if( d == frmFormatDate ) {
                if( c == okEditCommand ) {
                    currentFormat |= Result.FORMAT_DATE;
                    if( ((ChoiceGroup)frmFormatDate.get(0)).getSelectedIndex() == 0 )
                        currentFormat |= Result.FORMAT_DATE_DATE;
                    if( ((ChoiceGroup)frmFormatDate.get(1)).getSelectedIndex() == 0 )
                        currentFormat |= Result.FORMAT_DATE_TIME;
                    if( ((ChoiceGroup)frmFormatDate.get(2)).getSelectedIndex() == 1 )
                        currentFormat |= Result.FORMAT_DATE_YEAR4;
                    if( ((ChoiceGroup)frmFormatDate.get(3)).getSelectedIndex() == 1 )
                        currentFormat |= Result.FORMAT_DATE_ORDER;
                    if( ((ChoiceGroup)frmFormatDate.get(4)).getSelectedIndex() == 1 )
                        currentFormat |= Result.FORMAT_DATE_SEP;
                    currentCell.ll = currentFormat;
                    fitCell( cursorI, cursorJ );
                    sheet.changed = true;
                }
                frmFormatDate = null;
                setCurrent();
                return;
            }
            if( d == edit) {
                String msg = null;
                if( c == okEditCommand ) {
                    String formula = edit.getString();
                    try {
                        sheet.setFormula( cursorI, cursorJ, formula );
                        // refresh status info with the new value in INFO mode
                        if( statusInfo )
                            statusLine = getStatusInfo( cursorI, cursorJ );
                        else if( statusMem )
                            statusLine = getMemoryInfo();
                        sheet.calculateDepended( cursorI, cursorJ, true );
                        
                        fitCell( cursorI, cursorJ );
                        
                        sheet.changed = true;
                        
                    } 
                    catch( java.lang.ArithmeticException e ) {
                        msg = "Arithmetic error, divizion by zero or negative argument";
                    }
                    catch( Exception e ) {
                        msg = e.getMessage();
                    }
/*                    catch( ArithmeticException e ) {
                        msg = e.getMessage();
                    }*/
                }
                if( msg == null ) setCurrent();
                else warning( msg, edit );
                return;
            }
            
            if( c == Menu.cmdItem ) {
                int code = c.getPriority();
                if( code == CONFIRM_CANCEL )
                    return;
                if( code == CONFIRM_OK ) {
                    sheet.changed = false;
                    code = confirmCode;
                }
                switch( code ) {
                    case EXIT:
                        if( sheet.changed ) {
                            confirmCode = COMMAND_EXIT;
                            confirmMenu.start( parent, this );
                            break;
                        }
                        parent.notifyDestroyed();
                        break;
                    case CELL_COPY:
                        clipboardI = cursorI;
                        clipboardJ = cursorJ;
                        break;
                    case CELL_PASTE:
                        if( clipboardI >= 0 && clipboardJ >= 0 ) {
                            sheet.copyCell1( clipboardI, clipboardJ, cursorI-clipboardI, cursorJ-clipboardJ );
                        }
                        else
                            warning( "no cells selected" );
                        break;
                    case CELL_EDIT:
                        editCell( cursorI, cursorJ );
                        break;
                    case CELL_CLEAR:
                        sheet.clear( cursorI, cursorJ );
                        sheet.calculateDepended( cursorI, cursorJ, true );
                        break;
                    case FORMAT_NUMBER:
                        String prec = (currentFormat&Result.FORMAT_NUMBER)!=0 ? Integer.toString((currentFormat& 0x00F0) >> 4) : "";
                        
                        frmFormatNumber = new Form( "Number" );
                        
                        // 0
                        frmFormatNumber.append( new TextField( "# of digits after zero", prec, 2, TextField.NUMERIC ) );
                        
                        // 1
                        ChoiceGroup cg = new ChoiceGroup( "Thousands separator", Choice.EXCLUSIVE, new String[] { "Yes","No" }, null );
                        if( (currentFormat&Result.FORMAT_NUMBER_SEP) != 0 )
                            cg.setSelectedIndex(0,true);
                        else
                            cg.setSelectedIndex(1,true);
                        frmFormatNumber.append( cg );
                        
                        frmFormatNumber.addCommand( okEditCommand );
                        frmFormatNumber.addCommand( cancelEditCommand );
                        frmFormatNumber.setCommandListener( this );
                        display.setCurrent( frmFormatNumber );
                        break;
                    case FORMAT_DATETIME:
                        
                        frmFormatDate = new Form( "Datetime" );
                        
                        // 0
                        cg = new ChoiceGroup( "Display date", Choice.EXCLUSIVE, new String[] { "Yes","No" }, null );
                        if( (currentFormat&Result.FORMAT_DATE_DATE) != 0 )
                            cg.setSelectedIndex(0,true);
                        else
                            cg.setSelectedIndex(1,true);
                        frmFormatDate.append( cg );
                        
                        // 1
                        cg = new ChoiceGroup( "Display time", Choice.EXCLUSIVE, new String[] { "Yes","No" }, null );
                        if( (currentFormat&Result.FORMAT_DATE_TIME) != 0 )
                            cg.setSelectedIndex(0,true);
                        else
                            cg.setSelectedIndex(1,true);
                        frmFormatDate.append( cg );
                        
                        // 2
                        cg = new ChoiceGroup( "Number of digits in year", Choice.EXCLUSIVE, new String[] { "YY","YYYY" }, null );
                        if( (currentFormat&Result.FORMAT_DATE_YEAR4) != 0 )
                            cg.setSelectedIndex(1,true);
                        else
                            cg.setSelectedIndex(0,true);
                        frmFormatDate.append( cg );
                        
                        // 3
                        cg = new ChoiceGroup( "Month order", Choice.EXCLUSIVE, new String[] { "DD/MM","MM/DD" }, null );
                        if( (currentFormat&Result.FORMAT_DATE_ORDER) != 0 )
                            cg.setSelectedIndex(1,true);
                        else
                            cg.setSelectedIndex(0,true);
                        frmFormatDate.append( cg );
                        
                        // 6
                        cg = new ChoiceGroup( "Date separator", Choice.EXCLUSIVE, new String[] { "/ (slash)",". (dot)" }, null );
                        if( (currentFormat&Result.FORMAT_DATE_SEP) != 0 )
                            cg.setSelectedIndex(1,true);
                        else
                            cg.setSelectedIndex(0,true);
                        frmFormatDate.append( cg );
                        
                        frmFormatDate.addCommand( okEditCommand );
                        frmFormatDate.addCommand( cancelEditCommand );
                        frmFormatDate.setCommandListener( this );
                        display.setCurrent( frmFormatDate );
                        break;
/*
                case FORMAT_ALIGN:
                    formats = sheet.formats[cursorI][cursorJ];
 
                    frmFormatAlign = new Form( "Align" );
 
                    cg = new ChoiceGroup( "Horizontal", Choice.EXCLUSIVE, new String[] { "Default","Left","Right","Center" }, null );
                    if( (formats&Result.FORMAT_JUSTIFY_CENTER) == Result.FORMAT_JUSTIFY_CENTER )
                        cg.setSelectedIndex(3,true);
                    else if( (formats&Result.FORMAT_JUSTIFY_LEFT) != 0 )
                        cg.setSelectedIndex(1,true);
                    else if( (formats&Result.FORMAT_JUSTIFY_RIGHT) != 0 )
                        cg.setSelectedIndex(2,true);
                    else
                        cg.setSelectedIndex(0,true);
                    frmFormatAlign.append( cg );
 
                    frmFormatAlign.addCommand( okEditCommand );
                    frmFormatAlign.addCommand( cancelEditCommand );
                    frmFormatAlign.setCommandListener( this );
                    display.setCurrent( frmFormatAlign );
                    break;
 */
                    case FORMAT_CLEAR:
                        currentCell.ll = 0;
                        break;
                    case SHEET_NEW:
                    case SHEET_NEW1:
                        if( sheet.changed ) {
                            confirmCode = SHEET_NEW;
                            confirmMenu.start( parent, this );
                            return;
                        }
                        newSheet();
                        break;
                    case SHEET_DELETE:
                        if( sheet.changed ) {
                            confirmCode = SHEET_DELETE;
                            confirmMenu.start( parent, this );
                            return;
                        }
                        deleteSheet( sheet );
                        break;
                    case SHEET_LOAD:
                        if( sheet.changed ) {
                            confirmCode = SHEET_LOAD;
                            confirmMenu.start( parent, this );
                            return;
                        }
                        selectSheet();
                        break;
                    case SHEET_SAVE:
                        editName = new TextBox( "Sheet Name", sheet.name , 32-2, TextField.ANY );
                        editName.addCommand( okEditCommand );
                        editName.addCommand( cancelEditCommand );
                        editName.setCommandListener( this );
                        display.setCurrent( editName );
                        break;
                    case SHEET_RESIZE:
                        frmResize = new Form( "Sheet Size" );
                        frmResize.append( new TextField( "# of Rows", Integer.toString(sheet.rows), 3, TextField.NUMERIC ) );
                        frmResize.append( new TextField( "# of Columns", Integer.toString(sheet.columns), 3, TextField.NUMERIC ) );
                        frmResize.addCommand( okEditCommand );
                        frmResize.addCommand( cancelEditCommand );
                        frmResize.setCommandListener( this );
                        display.setCurrent( frmResize );
                        break;
                    case COLUMN_WIDTH:
                        editColumnWidth = new TextBox( "Column Width", Integer.toString( sheet.columnWidth[cursorJ] ), 3, TextField.NUMERIC );
                        editColumnWidth.addCommand( okEditCommand );
                        editColumnWidth.addCommand( cancelEditCommand );
                        editColumnWidth.setCommandListener( this );
                        display.setCurrent( editColumnWidth );
                        break;
                    case COLUMN_INSERT:
                        frmInsertColumn = new Form( "Insert Columns" );
                        frmInsertColumn.append( new TextField( "# of Columns", "1", 2, TextField.NUMERIC ) );
                        frmInsertColumn.append( new ChoiceGroup( "Resize", Choice.EXCLUSIVE, new String[] { "No","Yes" }, null ) );
                        frmInsertColumn.addCommand( okEditCommand );
                        frmInsertColumn.addCommand( cancelEditCommand );
                        frmInsertColumn.setCommandListener( this );
                        display.setCurrent( frmInsertColumn );
                        break;
                    case COLUMN_DELETE:
                        frmDeleteColumn = new Form( "Delete Columns" );
                        frmDeleteColumn.append( new TextField( "# of Columns", "1", 2, TextField.NUMERIC ) );
                        frmDeleteColumn.append( new ChoiceGroup( "Resize", Choice.EXCLUSIVE, new String[] { "No","Yes" }, null ) );
                        frmDeleteColumn.addCommand( okEditCommand );
                        frmDeleteColumn.addCommand( cancelEditCommand );
                        frmDeleteColumn.setCommandListener( this );
                        display.setCurrent( frmDeleteColumn );
                        break;
                    case ROW_INSERT:
                        frmInsertRow = new Form( "Insert Rows" );
                        frmInsertRow.append( new TextField( "# of Rows", "1", 2, TextField.NUMERIC ) );
                        frmInsertRow.append( new ChoiceGroup( "Resize", Choice.EXCLUSIVE, new String[] { "No","Yes" }, null ) );
                        frmInsertRow.append( new ChoiceGroup( "Copy values from row", Choice.EXCLUSIVE, new String[] { "No","above","below" }, null ) );
                        frmInsertRow.addCommand( okEditCommand );
                        frmInsertRow.addCommand( cancelEditCommand );
                        frmInsertRow.setCommandListener( this );
                        display.setCurrent( frmInsertRow );
                        break;
                    case ROW_DELETE:
                        frmDeleteRow = new Form( "Delete Rows" );
                        frmDeleteRow.append( new TextField( "# of Rows", "1", 2, TextField.NUMERIC ) );
                        frmDeleteRow.append( new ChoiceGroup( "Resize", Choice.EXCLUSIVE, new String[] { "No","Yes" }, null ) );
                        frmDeleteRow.addCommand( okEditCommand );
                        frmDeleteRow.addCommand( cancelEditCommand );
                        frmDeleteRow.setCommandListener( this );
                        display.setCurrent( frmDeleteRow );
                        break;
                    case ABOUT:
                        about();
                        break;
                }
            }
            
        }
        catch( RecordStoreFullException e ) {
            warning( "Cannot save the sheet - storage is full, try to delete old sheets");
        }
        catch( RecordStoreException e ) {
            warning( "Cannot load/save the sheet" );
        }
        catch( IOException e ) {
            warning( "Cannot load/save the sheet" );
        }
/*        
        catch( BadFormulaException e  ) {
            warning( "Cannot calculate the sheet" );
        }
 */
        catch( Exception e ) {
            warning( "Error: " + e.getMessage() );
        }
 
    }
    
    private void newSheet() {
        frmNewSheet = new Form( "New Sheet" );
        frmNewSheet.append( new TextField( "Name", Sheet.DEFAULT_NAME, 32-2, TextField.ANY ) );
        frmNewSheet.append( new TextField( "Rows", Integer.toString(Sheet.newRows), 3, TextField.NUMERIC ) );
        frmNewSheet.append( new TextField( "Columns", Integer.toString(Sheet.newColumns), 2, TextField.NUMERIC ) );
        frmNewSheet.addCommand( okEditCommand );
        frmNewSheet.addCommand( cancelEditCommand );
        frmNewSheet.setCommandListener( this );
        display.setCurrent( frmNewSheet );
    }
    
/*
    private Form makeBrowser( Result rr ) {
 
        frm = new Form( sheet.values[rr.i1][rr.j1] );
        for( int i=rr.j1; i<=rr.j2; i++ ) {
            Item
            frm.append
        }
    }
 */
    
    // doesn't repaints, btw
    /** fit columns width (to do: row height) to the current value
     * @see columnWidth
     */    
    private void fitCell( int i, int j ) {
        // fit the column width
        Result cell = sheet.getCell( i,j );
        Result value = cell.funcargs[0];
        if( value.type != Result.TYPE_STRING ) {
            String ss = value.toString(false,(int)cell.ll);
            int width = font.stringWidth( ss ) + 3;
            if( width > sheet.columnWidth[j] )
                sheet.columnWidth[j] = (short) width;
            sheet.changed = true;
        }
        // fit the height ...
    }
    
    private void editCell( int I, int J ) {
        
        Result cell = sheet.getCell( I,J );
        String formula = cell.str;
        
        if( !cell.hasFormula() ) formula = "=";
        
        edit = new TextBox( Result.cellAddress(I,J,0)+": Edit", formula, Sheet.MAX_FORMULA, TextField.ANY );
        
        edit.addCommand( okEditCommand );
        edit.addCommand( cancelEditCommand );
        edit.setCommandListener( this );
        
        display.setCurrent( edit );
    }
    
    private void about() {
        
        StringBuffer text = new StringBuffer(256);
        
        text.append("MicroCalc v");
        text.append( getVersion() );
        text.append( "\nFree memory " );
        text.append( getMemoryInfo() );
        text.append( '\n' );
        text.append( parent.ABOUT_MSG );
        text.append( '\n' );
        
        directions = new Alert("Help");
        directions.setTimeout(Alert.FOREVER);
        directions.setString( text.toString() );
        
        display.setCurrent( directions );
    }
    
    public static String getVersion() {
        String ss;
        ss = Integer.toString( Sheet.VERSION >> 16 );
        ss += "." + Integer.toString( (Sheet.VERSION >> 8) & 0xff );
        ss += "." + Integer.toString( Sheet.VERSION & 0xff );
        return ss;
    }

    public static List selectSheet(String title) {
        String stores[] = RecordStore.listRecordStores();
        List list = new List( title, List.IMPLICIT );
        if( stores != null ) {
            for( int i=0; i<stores.length; i++ ) {
                String ss = stores[i];
                if( ss.startsWith( Sheet.SHEET_PREFIX ) )
                    list.append( ss.substring( Sheet.SHEET_PREFIX.length() ), null );
            }
        }
        return list;
    }
    
    private void selectSheet() {
        List list = selectSheet( "Select sheet:" );
        list.addCommand( cancelEditCommand );
        list.addCommand( okEditCommand );
        list.addCommand( cmdDelete );
        list.setCommandListener( this );
        display.setCurrent( list );
        selectList = list;
    }
    
    public void warning( String msg, Displayable dd ) {
        Alert aa = new Alert( "Error:", msg, null, AlertType.ERROR );
        aa.setTimeout( Alert.FOREVER );
        display.setCurrent( aa, dd );
    }

    public void warning( String msg ) {
	warning( msg, this );
    }
   
    // todo - these functions shouldn't be here, were moved from Sheet class
    // must be separated from Canvas since it is not needed in Sync class
    // looks like that there must be Manager class, or MidletUtil
    public static void loadSheet( Sheet sheet, String _name )  throws IOException, BadFormulaException, RecordStoreException {
        RecordStore db = null;
        ByteArrayInputStream binp = null;
        
        try {
            db = RecordStore.openRecordStore( Sheet.SHEET_PREFIX + _name, false );
            // read header
            byte dd1[] = db.getRecord( 1 );
            byte dd2[] = db.getRecord( 2 );
            byte dd[] = new byte[dd1.length+dd2.length];
            System.arraycopy(dd1,0, dd,0, dd1.length);
            System.arraycopy(dd2,0, dd,dd1.length, dd2.length);
            binp = new ByteArrayInputStream( dd );
            sheet.loadSheet( _name, binp );
        }
        finally {
            if( db != null ) db.closeRecordStore();
            if( binp != null ) binp.close();
        }
        
    }

    public static void saveSheet( Sheet sheet ) throws RecordStoreException, IOException {
        
        if( !sheet.changed ) return;
        
        saveSheet( sheet, sheet.name );
        
        sheet.changed = false;
    }

    public static void saveSheet( Sheet sheet, String _name ) throws IOException, RecordStoreException {
        
        if( _name == null ) _name = Sheet.DEFAULT_NAME;
        
        // no need to save
        if( _name.compareTo(sheet.name) == 0 && !sheet.changed ) return;
        
        sheet.name = _name;
        
        try { RecordStore.deleteRecordStore( Sheet.SHEET_PREFIX + sheet.name ); } catch( RecordStoreNotFoundException e ) {}
        
        RecordStore db = null;
        ByteArrayOutputStream bout = null;
        
        try {
            db = RecordStore.openRecordStore( Sheet.SHEET_PREFIX + sheet.name, true );
            
            // write workbook header
            bout = new ByteArrayOutputStream();
            sheet.serverVersion++;
            sheet.saveSheet( bout );
            
            byte[] data = bout.toByteArray();
            db.addRecord( data, 0, 6 );
            db.addRecord( data, 6, data.length-6 );
        }
        finally {
            if( bout != null ) bout.close();
            if( db != null ) db.closeRecordStore();
        }
        
        sheet.changed = false;
    }

    public static void deleteSheet( String name )  {
        try { RecordStore.deleteRecordStore( Sheet.SHEET_PREFIX + name ); }
        catch( Exception e ) {}
    }
    
    public static void deleteSheet( Sheet sheet ) throws RecordStoreException {
        
        deleteSheet( sheet.name );
        
        sheet.rows = sheet.newRows; sheet.columns = sheet.newColumns;
        
        sheet.allocate();
        
        sheet.clearSheet();
        
        sheet.name = Sheet.DEFAULT_NAME;
        
    }

    void setCurrent() {
        display.setCurrent( this );
    }
    
    private void setStatusLine( String ss ) {
//        if( (ss == null && statusLine != null) || (ss != null && statusLine == null) || statusLine.compareTo( ss ) != 0 ) {
            statusLine = ss;
            statusLinePart = 0;
            repaint( 0, 0, sizeX, fontHeight );
//        }
    }    

    // assembles status line string in cell info mode
    // i,j - cell address
    private String getStatusInfo( int i, int j) {
        StringBuffer sb = new StringBuffer( Result.cellAddress(i,j,0) );
        sb.append( ": " );
        sb.append( sheet.getCell(i,j).str );
        return sb.toString();
    }
    
    private String getMemoryInfo() {
        StringBuffer sb = new StringBuffer(25);
        Runtime rt = Runtime.getRuntime();
        sb.append( Long.toString( rt.freeMemory() ) );
        sb.append( '[' );
        sb.append( Long.toString( rt.totalMemory() ) );
        sb.append( ']' );
        return  sb.toString();
    }
    
    private boolean intersect( 
        int x1, int y1, int x2, int y2, 
        int xx1, int yy1, int xx2, int yy2 ) {
            int xxx1 = x1 > xx1 ? x1 : xx1;
            int xxx2 = x2 < xx2 ? x2 : xx2;
            int yyy1 = y1 > yy1 ? y1 : yy1;
            int yyy2 = y2 < yy2 ? y2 : yy2;
            return xxx1 < xxx2 && yyy1 < yyy2;
    }
    
}
