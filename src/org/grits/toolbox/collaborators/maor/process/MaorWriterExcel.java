package org.grits.toolbox.collaborators.maor.process;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.grits.toolbox.ms.file.reader.impl.MzXmlReader;
import org.grits.toolbox.ms.om.data.Peak;
import org.grits.toolbox.ms.om.data.Scan;
import org.grits.toolbox.widgets.processDialog.ProgressDialog;
import org.grits.toolbox.widgets.progress.IProgressThreadHandler;

/**
 * Creates an Excel report for the MS data supplied by Maor.
 * 
 * @author D Brent Weatherly (dbrentw@uga.edu)
 *
 */
public class MaorWriterExcel {
	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorWriterExcel.class);

	public static DecimalFormat formatDec4 = new DecimalFormat("0.0000");
	public static DecimalFormat formatDec2 = new DecimalFormat("0.00");
	public static DecimalFormat formatDec1 = new DecimalFormat("0.0");
	protected Workbook m_objWorkbook = null;
	protected Sheet m_objSheet = null;
	protected int m_iRowCounter = 0;
	protected String m_outputFile = null;
	protected String m_inputFile = null;
	protected IProgressThreadHandler m_listener = null;
	protected MzXmlReader m_reader = null;
	protected int iSheetCount = 1;
	protected List<String> peakLabels = null;
	protected Map<Double, Map<String,Object>> peakToLabels = null;

	public final static int EXCEL_DEFAULT_COLUMN_WIDTH = 3500;
	public final static int EXCEL_MAX_NUM_ROWS = 20000;
	public final static double MZ_TOLERANCE = 0.2;
	protected CreationHelper createHelper = null;

	public void createNewFile(String _sInputFile, String _sOutputFile, 
			List<String> peakLabels, Map<Double, Map<String,Object>> peakToLabels,
			IProgressThreadHandler a_listener ) {
		this.m_objWorkbook = new XSSFWorkbook();
		this.m_listener = a_listener;
		this.m_outputFile = _sOutputFile;
		this.m_inputFile = _sInputFile;
		this.peakLabels = peakLabels;
		this.peakToLabels = peakToLabels;
		createHelper = m_objWorkbook.getCreationHelper();
	}

	/**
	 * Creates a new Excel sheet in the current workbook and writes the header line to it.
	 */
	public void createSheet() {
		this.m_objSheet = this.m_objWorkbook.createSheet( "Sheet" + this.iSheetCount);    
		this.m_iRowCounter = 0;
		writeHeadline();
		this.iSheetCount++;
	}

	/**
	 * Writes the current workbook to the specified output file. 
	 * 
	 * @throws IOException
	 */
	public void close() throws IOException {
		FileOutputStream t_fos = new FileOutputStream(this.m_outputFile);
		this.m_objWorkbook.write(t_fos);
		t_fos.close();
	}

	/**
	 * Increments the row counter so the current row won't be written to.
	 */
	public void writeEmptyLine() {
		this.m_iRowCounter++;
	}

	/**
	 * Writes the header row to the current Excel sheet in the current workbook
	 */
	public void writeHeadline() {
		Row t_row = this.m_objSheet.createRow(this.m_iRowCounter);
		int iColNum = 0;
		//		Cell t_cell = t_row.createCell(iColNum++);
		//		t_cell.setCellValue("name/m/z");
		//		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		//		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
		Cell t_cell = t_row.createCell(iColNum++);
		t_cell.setCellValue("min");
		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
		t_cell = t_row.createCell(iColNum++);
		t_cell.setCellValue("scan");
		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
		t_cell = t_row.createCell(iColNum++);
		t_cell.setCellValue("m/z");
		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
		t_cell = t_row.createCell(iColNum++);
		t_cell.setCellValue("TIC");
		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
		
		// add the column headers of any annotation information
		if( getPeakLabels() != null && ! getPeakLabels().isEmpty() ) {
			for( int i = 0; i < getPeakLabels().size(); i++ ) {
				String sLabel = getPeakLabels().get(i);
				t_cell = t_row.createCell(iColNum++);
				t_cell.setCellValue(sLabel);
				t_cell.setCellType(Cell.CELL_TYPE_STRING);
				this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);
			}
		}
		t_cell = t_row.createCell(iColNum++);
		t_cell.setCellValue("ms^2");
		t_cell.setCellType(Cell.CELL_TYPE_STRING);
		this.m_objSheet.setColumnWidth( iColNum, EXCEL_DEFAULT_COLUMN_WIDTH);

		writeEmptyLine();
	}

	/**
	 * Creates a new sheet in the Excel workbook when necessary
	 */
	protected void performPreWriteInits() {
		if( this.m_objSheet == null || this.m_iRowCounter == (EXCEL_MAX_NUM_ROWS-1) ) {
			createSheet();
		}
	}
	
	/**
	 * Main method to call in order to create the final annotated Excel report. It assumes the input file has been specified
	 * and the Excel Workbook and Sheet has been created.
	 */
	public void createReport() {
		this.m_reader = new MzXmlReader();
		List<Integer> lScans = this.m_reader.getScanList(m_inputFile, -1);
		int iMax = lScans.size();
		((ProgressDialog) m_listener).setMax(iMax);
		((ProgressDialog) m_listener).setProcessMessageLabel("Exporting data");
		int iCnt = 1;
		for( Integer iMS1 : lScans ) {
			if( ((ProgressDialog) m_listener).isCanceled() ) {
				break;
			}
			((ProgressDialog) m_listener).updateProgresBar("Scan: " + iCnt++ );
			List<Scan> scans = this.m_reader.readMzXmlFile(this.m_inputFile, 2, iMS1, -1);
			for( Scan scan : scans ) {
				if( ((ProgressDialog) m_listener).isCanceled() ) {
					break;
				}
				if( scan.getScanNo() == iMS1 ) {
					continue;
				}
				try {
					// determine centroid MS2 peaks and calculate TIC
					ScanSummary summary = new ScanSummary();
					summary.summarizeScan(scan);

					List<Object> lRow = new ArrayList();
					//					lRow.add("");
					double dRT = scan.getRetentionTime() / 60.0;
					lRow.add(Double.parseDouble( formatDec1.format( scan.getRetentionTime() / 60.0 ) ) );
					lRow.add(scan.getScanNo());
					if ( scan.getPrecursor() != null ) {
						lRow.add(Double.parseDouble( formatDec2.format(scan.getPrecursor().getPrecursorMz() ) ) );
					} else {
						lRow.add("");
					}
					lRow.add(Double.parseDouble( formatDec2.format(summary.getTIC())));
					
					// add to the Excel row any annotation information that is stored
					if( ! scan.getPeaklist().isEmpty() && getPeakLabels() != null && ! getPeakLabels().isEmpty() ) {
						// try to find an annotated theoretical peak based on precursor m/z
						Map<String,Object> peakLabels = getPeakLabel(scan.getPrecursor().getPrecursorMz());
						for( int i = 0; i < getPeakLabels().size(); i++ ) {
							String sLabel = getPeakLabels().get(i);
							if( peakLabels != null && peakLabels.containsKey(sLabel) ) {
								Object objVal = peakLabels.get(sLabel);
								lRow.add(objVal);
							} else {
								lRow.add("");
							}

						}
					}
					
					// summarize the MS2 peaks
					StringBuilder sb = new StringBuilder();
					for( int iCPeakInx = 0; iCPeakInx < summary.getCentroidPeaks().size(); iCPeakInx++ ) {
						if ( iCPeakInx > 0 ) {
							sb.append(", ");
						}
						Peak peak = scan.getPeaklist().get( summary.getCentroidPeaks().get(iCPeakInx) );
						double dMz = peak.getMz();
						sb.append(formatDec2.format(dMz));
						if( peak.getIntensity() == summary.getHighestPeak() ) {
							sb.append("^");
						}
					}
					lRow.add(sb.toString());
					writeRow(lRow);
				} catch ( Exception ex ) {
					;
				}
			}
		}
		((ProgressDialog) m_listener).setProcessMessageLabel("Done!");
	}

	/**
	 * Returns the mapping of column headers to values for the particular theoretical peak m/z
	 * 
	 * @param _dMz, the m/z of the peak from the annotation file
	 * @return
	 */
	protected Map<String,Object> getPeakLabel( double _dMz ) {
		if( getPeakToLabels() == null || getPeakToLabels().isEmpty() ) {
			return null;
		}
		Set<Double> peakKeys = getPeakToLabels().keySet();
		for( Double dKey : peakKeys ) {
			if( Math.abs( dKey - _dMz) < MZ_TOLERANCE ) { // need a constant for tolerance!
				return getPeakToLabels().get(dKey);
			}
		}
		return null;
	}

	/**
	 * Creates a new row in the current Excel sheet at the current position and then iterates over the objects in 
	 * the specified object vector to add those cells to the Excel row.
	 * 
	 * @param objValues, the object vector of source data
	 */
	public void writeRow( List<Object> objValues ) {
		performPreWriteInits();
		Row excelRow = this.m_objSheet.createRow(this.m_iRowCounter);
		for ( int iColNum = 0; iColNum < objValues.size(); iColNum++ ) {
			writeCell(excelRow, objValues, iColNum);
		}
		this.writeEmptyLine();    	
	}

	/**
	 * Populates the Excel row with the object in the table row vector at the position specified by the column number. 
	 * Depending on the data type of the object, attempts to set the cell type or style accordingly for Excel support.
	 * 
	 * @param _excelRow, the Excel row to populate
	 * @param _tableRow, the object vector of source data
	 * @param _iDataColNum, the column number of the cell
	 */
	protected void writeCell( Row _excelRow, List<Object> _tableRow, int _iDataColNum ) {
		Object oVal = _tableRow.get(_iDataColNum);
		if ( oVal == null ) 
			return;
		Cell t_cell = _excelRow.createCell(_iDataColNum);
		//		t_cell.setCellValue( oVal.toString() );
		if (oVal instanceof Number ) {
			if ( oVal instanceof Integer )
				t_cell.setCellValue((Integer) oVal);
			else
				t_cell.setCellValue(new Double(oVal.toString()));
			t_cell.setCellType(Cell.CELL_TYPE_NUMERIC);    				
		} else if ( oVal instanceof Boolean ) {
			//					t_cell.setCellValue( (Boolean) alRow.get(iColNum));   
			t_cell.setCellValue((Boolean) oVal);
			t_cell.setCellType(Cell.CELL_TYPE_BOOLEAN);    				
		} else if ( oVal instanceof Hyperlink ) {
			//cell style for hyperlinks
			//by default hyperlinks are blue and underlined
			CellStyle hlink_style = this.m_objWorkbook.createCellStyle();
			Font hlink_font = this.m_objWorkbook.createFont();
			hlink_font.setUnderline(Font.U_SINGLE);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			hlink_style.setFont(hlink_font);

			Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_URL);
			link.setAddress(((Hyperlink) oVal).getAddress());
			t_cell.setHyperlink(link);
			t_cell.setCellStyle(hlink_style);

			t_cell.setCellValue(((Hyperlink) oVal).getLabel());			
		} else {
			t_cell.setCellValue(oVal.toString());
			t_cell.setCellType(Cell.CELL_TYPE_STRING);    	
		}
		this.m_objSheet.setColumnWidth( _iDataColNum, EXCEL_DEFAULT_COLUMN_WIDTH);

	}

	/**
	 * @return
	 */
	public List<String> getPeakLabels() {
		return peakLabels;
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
	public Map<Double, Map<String, Object>> getPeakToLabels() {
		return peakToLabels;
	}

	/**
	 * @param peakToLabels
	 */
	public void setPeakToLabels(Map<Double, Map<String, Object>> peakToLabels) {
		this.peakToLabels = peakToLabels;
	}

	/**
	 * Utility class to process the MS2 peak list.
	 * 
	 * @author D Brent Weatherly (dbrentw@uga.edu)
	 *
	 */
	class ScanSummary {
		protected List<Integer> lCentroidPeaks = null;
		protected double dTIC = 0.0;
		protected double dHighestPeak = 0.0;
		
		/**
		 * Rudimentary method to determine list of centroided peaks from a scan. Starts at the first scan with intensity > 0 and 
		 * then determines the most intense peak w/in the MZ_TOLERANCE. That peak is the stored. It then continues until all scans 
		 * are processed.
		 * 
		 * @param scan, scan to process
		 * @return
		 */
		public void summarizeScan( Scan scan ) {
			lCentroidPeaks = new ArrayList<>();
			try {
				dTIC = 0.0;
				int iPeakInx = 0;
				dHighestPeak = 0.0;
				while( iPeakInx < scan.getPeaklist().size() ) {
					//		for( int iPeakInx = 0; iPeakInx < scan.getPeaklist().size(); iPeakInx++ ) {
					Peak curPeak = scan.getPeaklist().get(iPeakInx);
					while( curPeak.getIntensity() <= 0.0 ) {
						iPeakInx++;
					}
					// found base of peak;
					double dMaxInt = 0.0;
					int iMaxInx = iPeakInx;
					double dDelta = 0.0;
					while( iPeakInx < scan.getPeaklist().size() && 
							dDelta < MZ_TOLERANCE && 
							curPeak.getIntensity() > 0.0 ) {
						dTIC += curPeak.getIntensity();
						if( curPeak.getIntensity() > dMaxInt ) {
							dMaxInt = curPeak.getIntensity();
							iMaxInx = iPeakInx;
						}
						if( curPeak.getIntensity() > dHighestPeak ) {
							dHighestPeak = curPeak.getIntensity();
						}
						++iPeakInx;
						if( iPeakInx < scan.getPeaklist().size() ) {
							Peak nextPeak = scan.getPeaklist().get(iPeakInx);
							dDelta = nextPeak.getMz() - curPeak.getMz();
							curPeak = nextPeak;
						}
					}
					lCentroidPeaks.add(iMaxInx);
				}
			} catch( NullPointerException ex ) {
				logger.error(ex.getMessage(), ex);
			} catch( RuntimeException ex ) {
				logger.error(ex.getMessage(), ex);
			}
		}

		public double getTIC() {
			return dTIC;
		}
		
		public List<Integer> getCentroidPeaks() {
			return lCentroidPeaks;
		}
		
		public double getHighestPeak() {
			return dHighestPeak;
		}
	}
}