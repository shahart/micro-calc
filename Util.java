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

/**
 *
 * @author  mixa
 */
public abstract class Util {

    private Util() {
    }

    // converts data from text in hexs to binary
    public static byte[] convertAsciiToBin( String data ) {
        int len = data.length()/2;
        byte buf[] = new byte[len];
        for( int i=0; i<len; i++ ) {
            try {
                int bb1 = Integer.parseInt( data.substring(i*2,i*2+1), 16 );
                int bb2 = Integer.parseInt( data.substring(i*2+1,(i+1)*2), 16 );
                buf[i] = (byte) ( (bb1<<4) | bb2);
            }
            catch( NumberFormatException ex ) {}
        }
        return buf;
    }
    
    // converts data from text in hexs to binary
    public static String convertBinToAscii( byte[] data ) {
        int len = data.length;
        StringBuffer sb = new StringBuffer( len*2 );
        for( int i=0; i<data.length; i++ ) {
            byte bb = data[i];
            sb.append( Integer.toString( (bb>>4)&0xF, 16 ) );
            sb.append( Integer.toString( bb&0xF, 16 ) );
        }
        return sb.toString();
    }
    
}
