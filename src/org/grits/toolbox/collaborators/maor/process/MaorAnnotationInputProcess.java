package org.grits.toolbox.collaborators.maor.process;

import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.grits.toolbox.widgets.progress.CancelableThread;
import org.grits.toolbox.widgets.progress.IProgressThreadHandler;

/**
 * Thread class to read an Excel annotation file for Maor MS data.
 * 
 * @author D Brent Weatherly
 * @see MaorAnnotationInputReaderExcel
 */
public class MaorAnnotationInputProcess extends CancelableThread {

	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorAnnotationInputProcess.class);

	private String sInputFile = null;
	protected List<String> peakLabels = null;
	protected Map<Double, Map<String,Object>> peakToLabels = null;
	
	public MaorAnnotationInputProcess(String sInputFile) {
		this.sInputFile = sInputFile;
	}
	
	/**
	 * Creates an instance of the Excel annotation file reader
	 * 
	 * @param a_progressThreadHandler
	 * @return
	 */
	protected MaorAnnotationInputReaderExcel getNewMaorAnnotationInputReaderExcel(IProgressThreadHandler a_progressThreadHandler) {
		return new MaorAnnotationInputReaderExcel(getInputFile(), a_progressThreadHandler);
	}
	
	/** 
	 * Creates an instance of the Excel annotation file reader and processes it to determine the peak labels.
	 */
	public boolean threadStart(IProgressThreadHandler _progressThreadHandler) throws Exception{
		try{
			// write values to Excel
			MaorAnnotationInputReaderExcel excelReader = getNewMaorAnnotationInputReaderExcel(_progressThreadHandler);
			excelReader.readPeakAnnotationFile();
			setPeakLabels(excelReader.getPeakLabels());
			setPeakToLabels(excelReader.getPeakToLabels());
		}catch(Exception e)
		{
			logger.error(e.getMessage(), e);
			throw e;
		}
		return true;
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
