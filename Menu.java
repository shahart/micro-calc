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

public final class Menu implements CommandListener {

    // I don't use classes to save heap
    // however it would be better with MenuItem class to inherit from it prompts 
    // and edit boxes

    static final short MASK_CONFIRM	= 0x0100;
    static final short MASK_PROMPT	= 0x0200;	// don't restore old screen
    static final short MASK_SELECTED	= 0x0400;

    String labels[];
    short codes[];				// to generate result command. if == 0 then it's a submenu
    byte parents[];

    int current;
    int nselected;				// selected item on the current level

    List list;

    public static Command cmdItem;

    private static final Command cmdOK		= new Command("OK", Command.OK, 1);
    private static final Command cmdBack	= new Command("Back", Command.BACK, 2);
    private static final Command cmdCancel	= new Command("Cancel", Command.SCREEN, 3);

    private Form frmConfirm;

    private CommandListener cmdHandler;
    private MIDlet midlet;
    private Displayable oldScreen;
    private Display display;

    public Menu( String _labels[], short _codes[], byte _parents[] ) {

	labels = _labels;
	codes = _codes;
	parents = _parents;

    }

    public void start( MIDlet _midlet, CommandListener _cmdHandler, int start, Displayable _oldScreen ) {
	midlet = _midlet;
	cmdHandler = _cmdHandler;
	display = Display.getDisplay( midlet );
	oldScreen = _oldScreen;
	if( oldScreen == null ) 
	    oldScreen = display.getCurrent();
	// root must be ITEM_MENU
	loadMenu( start );
    }

    public void start( MIDlet _midlet, CommandListener _cmdHandler ) {
	start( _midlet, _cmdHandler, 0, null );
    }

    private int menuItems( int nn ) {
	int nnn = 0;
	for( int i=0; i<labels.length; i++ )
	    if( parents[i] == nn ) nnn++;
	return nnn;
    }

    private int nextItem( int parent, int start ) {
	start++;
	while( start<labels.length ) {
	    if( parents[start] == parent ) 
		return start;
	    start++;
	}
	return -1;
    }

    private int itemIndex( int parent, int number ) {
	for( int i=0; i<labels.length; i++ ) {
	    if( parents[i] == parent ) number--;
	    if( number < 0 ) return i;
	}
	return -1;
    }

    private String getPath( int nn ) {
	if( nn == 0 ) return "";
	if( parents[nn] == 0 ) return labels[nn];
	return getPath( parents[nn] ) + "/" + labels[nn];
    }

    private void loadMenu( int start ) {

	int j;

	int length = menuItems( start );

	String title = start == 0 ? labels[start] : getPath(start);

	list = new List( title, List.IMPLICIT );

	j = start;
	nselected = 0;
	int nn = 0;
	for( int i=0; i<length; i++ ) {
	    j = nextItem( start, j );
	    list.append( labels[j], null );
	    if( (codes[j] & MASK_SELECTED) != 0 ) {
		nselected = j;
		nn = i;
	    }
	}

	list.setSelectedIndex( nn, true );

	current = start;

	list.addCommand( cmdCancel );
	list.addCommand( cmdBack );
	list.addCommand( cmdOK );
	list.setCommandListener( this );

	display.setCurrent( list );

    }

    public void commandAction( Command cc, Displayable dd ) {

      if( dd == list ) {

	if( cc == List.SELECT_COMMAND || cc == cmdOK ) {
	    int nn = list.getSelectedIndex();
	    if( nn < 0 ) return;
	    current = itemIndex( current, nn );
	    codes[current] |= MASK_SELECTED;
	    if( (codes[current]&~MASK_SELECTED) == 0 ) {
		loadMenu( current );
		return;
	    }
	    // TYPE_ITEM
	    codes[nselected] &= ~MASK_SELECTED;
	    if( (codes[current] & MASK_CONFIRM) != 0 ) {
		frmConfirm = new Form( "Confirmation" );
	   	frmConfirm.append( "Please confirm operation '" + getPath(current) + "'" );
		frmConfirm.addCommand( cmdOK );
		frmConfirm.addCommand( cmdBack );
	 	frmConfirm.setCommandListener( this );
		display.setCurrent( frmConfirm );
		return;
	    }

	    executeCommand( current );
	    current = parents[current];

	}
	else if( cc == cmdCancel ) {
	    restoreScreen();
	}
	else if( cc == cmdBack ) {
	    if( current == 0 )
		restoreScreen();		// top menu, close menu
	    else
	    	loadMenu( parents[current] );
	}
   	return;
      }

      if( dd == frmConfirm ) {

      	if( cc == cmdOK )
      	    executeCommand( current );
      	else
	    display.setCurrent( list );

	current = parents[current];

      }

    }

    private void restoreScreen() {
	if( oldScreen != null ) display.setCurrent( oldScreen );
    }

    private void executeCommand( int nn ) {
	cmdItem = new Command( "", Command.SCREEN, codes[nn] & ~(MASK_CONFIRM | MASK_PROMPT | MASK_SELECTED) );
	cmdHandler.commandAction( cmdItem, list );
	// if there will be a prompt don't restore the screen
	if( ( codes[nn] & MASK_PROMPT ) == 0 ) 
	    restoreScreen();
    	frmConfirm = null;
//	list = null;
    }

}
