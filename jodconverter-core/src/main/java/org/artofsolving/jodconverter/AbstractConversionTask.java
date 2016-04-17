//
// JODConverter - Java OpenDocument Converter
// Copyright 2009 Art of Solving Ltd
// Copyright 2004-2009 Mirko Nasato
//
// JODConverter is free software: you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public License
// as published by the Free Software Foundation, either version 3 of
// the License, or (at your option) any later version.
//
// JODConverter is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General
// Public License along with JODConverter.  If not, see
// <http://www.gnu.org/licenses/>.
//
package org.artofsolving.jodconverter;

import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;
import static org.artofsolving.jodconverter.office.OfficeUtils.cast;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUnoProperties;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUrl;

import java.io.File;
import java.util.Map;

import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeTask;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XPrintAreas;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.task.ErrorCodeIOException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XRefreshable;

/**
 * Added an overridable method to handle processing of the document before it
 * get's converted
 * 
 * @author <a href="mailto:tdelprat@nuxeo.com">Tiry</a>
 * 
 */
public abstract class AbstractConversionTask implements OfficeTask {

    private final File inputFile;
    private final File outputFile;
    private final boolean isFix;

    // isFix = true then if spread sheet set print area
    
    public AbstractConversionTask(File inputFile, File outputFile,boolean isFix) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
        this.isFix=isFix;
    }

    protected abstract Map<String, ?> getLoadProperties(File inputFile);

    protected abstract Map<String, ?> getStoreProperties(File outputFile,
            XComponent document);

    public void execute(OfficeContext context) throws OfficeException {
        XComponent document = null;
        try {
            document = loadDocument(context, inputFile);
            storeDocument(document, outputFile);
        } catch (OfficeException officeException) {
            throw officeException;
        } catch (Exception exception) {
            throw new OfficeException("conversion failed", exception);
        } finally {
            if (document != null) {
                XCloseable closeable = cast(XCloseable.class, document);
                if (closeable != null) {
                    try {
                        closeable.close(true);
                    } catch (CloseVetoException closeVetoException) {
                        // whoever raised the veto should close the document
                    }
                } else {
                    document.dispose();
                }
            }
        }
    }

    protected XComponent loadDocument(OfficeContext context, File inputFile)
            throws OfficeException {
        if (!inputFile.exists()) {
            throw new OfficeException("input document not found");
        }
        Object desktopService = context.getService(SERVICE_DESKTOP);
        XComponentLoader loader = cast(XComponentLoader.class, desktopService);

        Map<String, ?> loadProperties = getLoadProperties(inputFile);
        XComponent document = null;
        try {
            document = loader.loadComponentFromURL(toUrl(inputFile), "_blank",
                    0, toUnoProperties(loadProperties));
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load document: "
                    + inputFile.getName(), illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load document: "
                    + inputFile.getName() + "; errorCode: "
                    + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not load document: "
                    + inputFile.getName(), ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load document: "
                    + inputFile.getName());
        }

        handleDocumentLoaded(document);
        //-----------------------------------------------------------------------
        // FIX CONVERTION CRASH ON BIG EXCEL DATA WITHOUT PRINT AREA OR WITH BIG PRINT AREA!
        //-----------------------------------------------------------------------
        XSpreadsheetDocument sx = cast(XSpreadsheetDocument.class, document);
        if (sx!= null && isFix) 
        {
            String n1 = sx.getSheets().getElementNames()[0];
            XSpreadsheet s=null;
    		try 
    		{
    			s = cast(XSpreadsheet.class, sx.getSheets().getByName(n1));
    	        XSheetCellCursor cursor = s.createCursor();
    	        XUsedAreaCursor cursor2 = (XUsedAreaCursor) UnoRuntime.queryInterface(XUsedAreaCursor.class, cursor);
    	        cursor2.gotoStartOfUsedArea(false);
    	        cursor2.gotoEndOfUsedArea(true);
    	        XCellRangeAddressable lAddressable = (XCellRangeAddressable) UnoRuntime.queryInterface(XCellRangeAddressable.class, cursor2);
    	        if (lAddressable.getRangeAddress().EndRow > 3000) 
    	        {
        	        XPrintAreas printAreas = (XPrintAreas)UnoRuntime.queryInterface(XPrintAreas.class, s);
        	        CellRangeAddress[] crr = printAreas.getPrintAreas();
        	        if (crr == null || crr.length == 0) 
        	        {
        	        	CellRangeAddress e = lAddressable.getRangeAddress(); 
        	        	e.EndRow=3000;        	        
            	        printAreas.setPrintAreas(new CellRangeAddress[] {e});
        	        } else {
        	        	boolean ok=false;
        	        	for (CellRangeAddress e : crr) 
        	        	{
        	        		if (e.EndRow - e.StartRow > 3000) {
        	        			e.EndRow = e.StartRow + 3000;
        	        			ok=true;
        	        		}
        	        	}
        	        	if (ok)
        	        		printAreas.setPrintAreas(crr);
        	        }
    	        }
    	        //-----------------------------------------------------------------------
    		} catch (NoSuchElementException | WrappedTargetException e) {}
        }
        //-----------------------------------------------------------------------
        return document;
    }

    // make this processing overridable by child classes
    protected void handleDocumentLoaded(XComponent document) {
        XRefreshable refreshable = cast(XRefreshable.class, document);
        if (refreshable != null) {
            refreshable.refresh();
        }
    }

    protected void storeDocument(XComponent document, File outputFile)
            throws OfficeException {
        Map<String, ?> storeProperties = getStoreProperties(outputFile,
                document);
        if (storeProperties == null) {
            throw new OfficeException("unsupported conversion");
        }
        try {
            cast(XStorable.class, document).storeToURL(toUrl(outputFile),
                    toUnoProperties(storeProperties));
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not store document: "
                    + outputFile.getName() + "; errorCode: "
                    + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not store document: "
                    + outputFile.getName(), ioException);
        }
    }

}
