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

import javax.microedition.midlet.*;
import javax.microedition.lcdui.*;
import javax.microedition.rms.*;
import java.util.*;

public final class MicroCalc extends MIDlet {

    // platform specific things are added here to save space (less classes)
    // TODO - define user interface here (specific for different phones)
    // TODO - add FullCanvas here?
    public static final short DEFAULT_PRECISION = MathFP._digits-3;	// different on sl45, a bug in J2ME long mul operation
    static final String ABOUT_MSG =
	"Use '=' before formulas. " +
	"Operators: + - / * ^(power) &(string concatenation), logical operators: = != > < >= <=\n" +
    	"Functions: EXP, LN, SQRT, ABS, POWER, SUM, IF, AND, OR, RAND, MAX, MIN, PMT, SUMSQ, COUNT, AVERAGE, STDEV\n" +
    	"Data types: boolean, long, decimal (fixed point 40/24 bits), string, datetime.\n" +
    	"Press <#><0> for NUMERIC mode, then <#> to finish it.\n<#><9> - on/off cell info mode.\n<#><1> - copy cell<#><2> - paste.\n" +
            "modified by me (Shahar) to make ex/import work plus additional functions.\n" +
            "For full instructions, news and latest updates visit us at www.wapindustrial.com (wap page http://www.wapindustrial.com/microcalc/microcalc.wml).\n" +
            "MicroCalc package includes 'Sync' utility to save sheets on network server to your personal folder.\n" +
            "Midlet is developed by WAP INDUSTRIAL [support@wapindustrial.com]."; 
    
    MicroCalcCanvas canvas;

    public MicroCalc() {
    }

    // pseudo interface to show progress when loading/saving sheets
    public void loadSheet( String name ) throws java.io.IOException,BadFormulaException,javax.microedition.rms.RecordStoreException {
        canvas.display.setCurrent( new TempScreen( canvas, name, canvas, canvas.selectList ) );
    }

    public void saveSheet( String name ) throws java.io.IOException,javax.microedition.rms.RecordStoreException {
	// temscreen knows what to load (direct access to canvas ) ugly
        canvas.display.setCurrent( new TempScreen( canvas, null, canvas, canvas.editName ) );
    }

    protected void destroyApp( boolean  b) {
//	try { canvas.sheet.saveSheet();	}
//	catch( Exception e ) {}			// do not report if we can't save the sheet
    }

    protected void startApp() {
	canvas = new MicroCalcCanvas( this );
//	canvas.setWindow( 0, 0, MicroCalcCanvas.CURSOR_DOWN, MicroCalcCanvas.CURSOR_RIGHT );
	MicroCalcCanvas.startMenu.start( this, canvas, 0,  canvas );
    }

    protected void pauseApp( ) {
    }

}

