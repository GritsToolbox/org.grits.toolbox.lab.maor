package org.grits.toolbox.collaborators.maor.process;

import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.grits.toolbox.widgets.progress.CancelableThread;
import org.grits.toolbox.widgets.progress.IProgressThreadHandler;

/**
 * Thread class to create an Excel report for Maor MS data.
 * 
 * @author D Brent Weatherly
 * @see MaorWriterExcel
 */
public class MaorExportProcess extends CancelableThread {

	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorExportProcess.class);

	private String sOutputFile = null;
	private String sInputFile = null;
	protected List<String> peakLabels = null;
	protected Map<Double, Map<String,Object>> peakToLabels = null;
	
	public MaorExportProcess( List<String> peakLabels, Map<Double, Map<String,Object>> peakToLabels ) {
		this.peakLabels = peakLabels;
		this.peakToLabels = peakToLabels;
	}
	
	protected MaorWriterExcel getNewMSAnnotationWriterExcel() {
		return new MaorWriterExcel();
	}
	
	/** 
	 * Creates an instance of the Excel writer, creates the report, and writes the report to file.
	 */
	@Override
	public boolean threadStart(IProgressThreadHandler _progressThreadHandler) throws Exception{
		try{
			// write values to Excel
			MaorWriterExcel writerExcel = getNewMSAnnotationWriterExcel();

			writerExcel.createNewFile(getInputFile(), getOutputFile(), 
					getPeakLabels(), getPeakToLabels(), _progressThreadHandler);	
			writerExcel.createSheet();
			writerExcel.createReport();
			writerExcel.close();
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
			throw e;
		}
		return true;
	}

	/**
	 * @return
	 */
	public String getOutputFile() {
		return sOutputFile;
	}

	/**
	 * @param _sOutputFile
	 */
	public void setOutputFile(String _sOutputFile) {
		this.sOutputFile = _sOutputFile;
	}
	
	/**
	 * @return
	 */
	public String getInputFile() {
		return sInputFile;
	}
	
	/**
	 * @param sInputFile
	 */
	public void setInputFile(String sInputFile) {
		this.sInputFile = sInputFile;
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
}
