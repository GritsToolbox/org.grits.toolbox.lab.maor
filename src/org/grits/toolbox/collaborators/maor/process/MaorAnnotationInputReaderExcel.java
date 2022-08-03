package org.grits.toolbox.collaborators.maor.process;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.grits.toolbox.widgets.progress.IProgressThreadHandler;

/**
 * Processes an Excel file containing information to annotate peaks of interest. 
 * It makes now assumptions about how much information can be used to annotate peaks, except for these rules:
 * 1. The first row of the file contains the column headers
 * 2. The following rows contain the information for each peak
 * 3. The column header for the theoretical peak m/z must be "m/z"
 * 4. Reading stops after first empty column. Any remaining columns will be ignored
 * Column header of most recent file:
 *   ID		m/z		Name	Formula		CAS		KEGG
 *  
 * @author D Brent Weatherly (dbrentw@uga.edu)
 *
 */
public class MaorAnnotationInputReaderExcel {
	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorAnnotationInputReaderExcel.class);
	protected Workbook objWorkbook = null;
	protected Sheet objSheet = null;
	protected String inputFile = null;
	protected IProgressThreadHandler listener = null;
	protected int iMassIndex = -1;

	protected List<String> peakLabels = null;
	// create a map of the m/z value of peak to a map of peak header names to the values for that peak
	protected Map<Double, Map<String,Object>> peakToLabels = null;

	public MaorAnnotationInputReaderExcel(String _sInputFile, IProgressThreadHandler a_listener ) {
		this.listener = a_listener;
		this.inputFile = _sInputFile;
		setPeakLabels(new ArrayList<>());
		setPeakToLabels(new HashMap<>());
	}

	/**
	 * Opens the Excel file and reads it to get the annotation information for peaks
	 * of interest. The method makes no assumption about how many columns are present but 
	 * the must be contiguous (reader stops when empty column header is encountered)
	 */
	public void readPeakAnnotationFile() {
		try {
			this.objWorkbook = openInputFile();
			this.objSheet = getWorkbook().getSheetAt(getWorkbook().getActiveSheetIndex());
			readHeader();
			populatePeakLabels();
			close();
		} catch (NullPointerException npe ) {
			logger.error(npe.getMessage(), npe);
		} catch( RuntimeException rte ) {
			logger.error(rte.getMessage(), rte);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}

	}

	/**
	 * Assumes the workbook is open and the sheet is set. It then reads the first line of the 
	 * sheet, assuming that is the header line. The method makes no assumption about 
	 * how many columns are present but they must be contiguous 
	 * (reader stops when empty column header is encountered).
	 * Sets the value of iMassIndex to store the column number of the peak mass
	 */
	protected boolean readHeader() {
		try {		
			// assuming the header row is the first row
			iMassIndex = -1;
			Row row = this.objSheet.getRow(0);
			if( getPeakLabels() == null ) {
				setPeakLabels( new ArrayList<String>() );
			} else {
				getPeakLabels().clear();
			}
			int iColNum = 0;
			Cell cell = null;
			String sHeader = null; 
			do {
				cell = row.getCell(iColNum++);
				sHeader = null;
				if( cell != null ) {
					sHeader = cell.getStringCellValue().trim();
					if( ! sHeader.equals("") ) {
						if( sHeader.equalsIgnoreCase("m/z") ) {
							iMassIndex = getPeakLabels().size();
						}
						getPeakLabels().add(sHeader);
					}
				}
			} while( cell != null && sHeader != null && ! sHeader.equals(""));
			if( iMassIndex >= 0 ) {
				return true;
			}
		} catch (NullPointerException npe ) {
			logger.error(npe.getMessage(), npe);
		} catch( RuntimeException rte ) {
			logger.error(rte.getMessage(), rte);
		}
		return false;
	}

	/**
	 * Assumes the workbook is open and the sheet is set. It assumes the list of column headers
	 * has been read (readHeader). It then reads, starting at the second line, to the end of the file 
	 */
	protected boolean populatePeakLabels() {
		try {	
			assert this.iMassIndex >= 0;
			int iRowInx = 1;
			Row row = null;
			do {
				row = this.objSheet.getRow(iRowInx++);
				Double dPeakVal = null;
				if( row != null ) {
					Map<String, Object> mPeakLabels = new HashMap<>();
					for(int i = 0; i < getPeakLabels().size(); i++ ) {
						Cell cell = row.getCell(i);
						Object objVal = null;
						Hyperlink hLink = null;
						try {
							hLink = cell.getHyperlink();
						} catch( Exception e ) {
							objVal = null;
						}
						if( objVal == null ) {
							try {
								objVal = cell.getNumericCellValue();
								double dVal = (Double) objVal;
								int iVal = (int) dVal;
								if( (double) iVal == dVal ) { // this was an integer value!
									objVal = iVal;
								}
							} catch( Exception e ) {
								objVal = null;
							}
						}
						if( objVal == null ) {
							try {
								objVal = cell.getBooleanCellValue();
							} catch( Exception e ) {
								objVal = null;
							}							
						}
						if( objVal == null ) {
							try {
								objVal = cell.getStringCellValue();
							} catch( Exception e ) {
								objVal = null;
							}							
						}
						if( i == this.iMassIndex && objVal != null ) {
							try {
								dPeakVal = (Double) objVal;
							} catch( Exception e ) {
								dPeakVal = null;
							}							
						}
						
						// if the cell is a hyperlink, we need to store the label too, but it seems the label 
						// is the same as the address, so set the label as the current value, so we can write the 
						// label in the final report
						if( hLink != null ) {
							String sLabel = objVal.toString();
							hLink.setLabel(sLabel);
							objVal = hLink;
						}
						mPeakLabels.put(getPeakLabels().get(i), objVal);
					}				
					if( dPeakVal != null ) {
						getPeakToLabels().put(dPeakVal, mPeakLabels);
					}
				}
			} while( row != null && iRowInx <= this.objSheet.getLastRowNum() );
			return true;
		} catch (NullPointerException npe ) {
			logger.error(npe.getMessage(), npe);
		} catch( RuntimeException rte ) {
			logger.error(rte.getMessage(), rte);
		}
		return false;
	}

	/**
	 * Creates a new Excel workbook using the specified input file.
	 * 
	 * @return
	 * @see XSSFWorkbook
	 */
	protected XSSFWorkbook openInputFile() {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(getInputFile());
			return workbook;
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
		return null;
	}

	/**
	 * Closes the workbook.
	 * 
	 * @throws IOException
	 */
	public void close() throws IOException {
		getWorkbook().close();
	}

	/**
	 * @param peakLabels
	 */
	public void setPeakLabels(List<String> peakLabels) {
		this.peakLabels = peakLabels;
	}

	/**
	 * @return
	 */
	public List<String> getPeakLabels() {
		return peakLabels;
	}

	/**
	 * @param _dMz
	 * @return
	 */
	protected Map<String, Object> getPeakLabels( double _dMz ) {
		if( getPeakToLabels() == null || getPeakToLabels().isEmpty() || ! getPeakToLabels().containsKey(_dMz) ) {
			return null;
		}
		return getPeakToLabels().get(_dMz);
	}
	
	/**
	 * @param peakToLabels
	 */
	public void setPeakToLabels(Map<Double, Map<String, Object>> peakToLabels) {
		this.peakToLabels = peakToLabels;
	}

	/**
	 * @return
	 */
	public Map<Double, Map<String, Object>> getPeakToLabels() {
		return peakToLabels;
	}

	/**
	 * @return
	 */
	protected String getInputFile() {
		return inputFile;
	}

	/**
	 * @return
	 */
	protected Sheet getCurrentSheet() {
		return objSheet;
	}

	/**
	 * @return
	 */
	protected Workbook getWorkbook() {
		return objWorkbook;
	}

}