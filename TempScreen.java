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

import javax.microedition.lcdui.*;

final class TempScreen extends Canvas implements Runnable {
    private String sheetname;
    private Displayable okScreen, errScreen;
    private MicroCalcCanvas canvas;
    private boolean done;

    TempScreen( MicroCalcCanvas _canvas, String _sheetname, Displayable _okScreen, Displayable _errScreen ) { 
        sheetname = _sheetname;
        okScreen = _okScreen;
        errScreen = _errScreen;
        canvas = _canvas;
        done = false;
    }
    
    protected void paint(Graphics g) {
        g.setColor( 0xFFFFFF );
        g.fillRect( 0,0,getWidth(),getHeight() );
        g.setColor( 0 );
        g.drawString(sheetname != null ? "Loading..." : "Saving...", getWidth() / 2, getHeight() / 2, Graphics.TOP | Graphics.HCENTER);
        canvas.display.callSerially(this);
    }

    public void run() {
        if( done) return;
        done = true;
        try {
            if( sheetname == null ) {
                canvas.saveSheet( canvas.sheet );
                canvas.display.setCurrent( okScreen );
            }
            else {
                canvas.loadSheet( canvas.sheet, sheetname );
                canvas.setWindow( 0, 0, MicroCalcCanvas.CURSOR_DOWN, MicroCalcCanvas.CURSOR_RIGHT );
                canvas.display.setCurrent( okScreen );
            }
        }
        catch( Exception e ) {
            Alert err = new Alert("Cannot " + (sheetname == null ? "save" : "load") + " the sheet, error: ", e.getMessage(), null, AlertType.ERROR);
            err.setTimeout( Alert.FOREVER );
            canvas.display.setCurrent(err, errScreen);
        }
    }
}
