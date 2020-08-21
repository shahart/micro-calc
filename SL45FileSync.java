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

import javax.microedition.io.file.FileConnection;
import javax.microedition.lcdui.*;
import javax.microedition.midlet.*;
import javax.microedition.rms.*;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.Enumeration;
import javax.microedition.io.Connector;

import com.sun.midp.io.j2me.storage.File;
import com.wapindustrial.calc.*;

public final class SL45FileSync extends MIDlet implements CommandListener {

    private static final Command cmdExit	= new Command("Exit", Command.EXIT, 2);
    private static final Command cmdOK 		= new Command("OK", Command.OK, 1);

    List selectList;

    Display display;

    public SL45FileSync() {
	display = Display.getDisplay( this );
	selectSheet();
    }

    protected void destroyApp( boolean  b) {
    }

    protected void startApp() {
    }

    protected void pauseApp( ) {
    }

    private void selectSheet() {
	String stores[] = RecordStore.listRecordStores();
	selectList = new List( "Export sheet:", List.IMPLICIT );
	if( stores != null ) {
	    for( int i=0; i<stores.length; i++ ) {
	    	String ss = stores[i];
	    	if( ss.startsWith( Sheet.SHEET_PREFIX ) )
		    selectList.append( ss.substring( Sheet.SHEET_PREFIX.length() ), null );
	    }
	}
	selectList.addCommand( cmdExit );
	selectList.addCommand( cmdOK );
	selectList.setCommandListener( this );
        display.setCurrent( selectList );
    }

    public void commandAction( Command c, Displayable d ) {

	if( c == cmdExit ) {
	    notifyDestroyed();
	    return;
	}

	if( d == selectList ) {
	    if( c == List.SELECT_COMMAND || c == cmdOK  ) {
	    	int nn = selectList.getSelectedIndex();
	    	if( nn >= 0 ) {
	    	    final String fname = selectList.getString( nn );
            	    showWait("Saving...", 
			new Runnable()
              		    { 
				public void run() { 
				    actualSave(fname); 
				} 
			    }
		    );
	    	}
	    }
	    else
		show();
	    return;
	}

    }

    private static void sheetToSYLK( Sheet sheet ) throws IOException {

        // TODO?getAppProperty("your-file-location");
        FileConnection localFileConnection1 = (FileConnection)Connector.open("file://localhost/Others/" + sheet.name + ".slk.txt");
        if (localFileConnection1.exists())
          localFileConnection1.delete();
        localFileConnection1.create();
        Object localObject1 = localFileConnection1.openDataOutputStream();
        String sss = sheet.toSylk();
        byte[] arrayOfByte = sss.getBytes(/*"UTF8"*/);
        ((DataOutputStream)localObject1).write(arrayOfByte,0,arrayOfByte.length);
        ((DataOutputStream)localObject1).close();
        localFileConnection1.close();
    }

    /** Show a wait message. Don't forget to set the current displayable back. */
    private void showWait(final String message, final Runnable action) {
    	Canvas wait = new Canvas() {
            protected void paint(Graphics g) {
          	g.drawString(message, getWidth() / 2, getHeight() / 2, Graphics.TOP | Graphics.HCENTER);
          	display.callSerially(action);
            }
      	};
    	display.setCurrent(wait);
    }

    /** Internal save function */
    private void actualSave(String fname) {
    	try {
	    Sheet sheet;
            sheet = new Sheet( Sheet.newRows, Sheet.newColumns, 0, 0 );
    	    MicroCalcCanvas.loadSheet( sheet, fname );
    	    sheetToSYLK( sheet );
            show();
	}
    	catch( Exception e ) { 
            Alert err = new Alert("Error: ", e.getMessage(), null, AlertType.ERROR);
            display.setCurrent(err, selectList);
    	}
    }

    private void show() {
    	display.setCurrent(selectList);
    	selectList.setCommandListener(this);
    }

}
