package org.grits.toolbox.collaborators.maor.adaptor;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.apache.log4j.Logger;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Shell;
import org.grits.toolbox.collaborators.maor.dialog.MaorMSExportDialog.ExportTypes;
import org.grits.toolbox.collaborators.maor.process.MaorAnnotationInputProcess;
import org.grits.toolbox.collaborators.maor.process.MaorAnnotationInputReaderExcel;
import org.grits.toolbox.core.datamodel.Entry;
import org.grits.toolbox.core.utilShare.ErrorUtils;
import org.grits.toolbox.entry.ms.property.MassSpecEntityProperty;
import org.grits.toolbox.entry.ms.property.MassSpecProperty;
import org.grits.toolbox.widgets.processDialog.ProgressDialog;

/**
 * Adapter to support reading files containing annotation information for MS data.
 * Currently only supports Excel version.
 * @author D Brent Weatherly (dbrentw@uga.edu)
 * @see MaorAnnotationInputReaderExcel
 *
 */
public class MaorAnnotationFileAdapter extends SelectionAdapter {

	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorAnnotationFileAdapter.class);

	protected Shell shell = null;
	protected String fileExtension = null;
	protected Entry msEntry = null;
	protected ExportTypes exportType;
	protected String sInputFile = null;
	protected List<String> peakLabels = null;
	protected Map<Double, Map<String,Object>> peakToLabels = null;

	protected MassSpecProperty getProperty() {
		MassSpecEntityProperty entityProp = (MassSpecEntityProperty) this.msEntry.getProperty();
		MassSpecProperty property = (MassSpecProperty) entityProp.getParentProperty();
		return property;
	}

	/**
	 * This is needed because of the way we create dialog titles
	 * 
	 * @param _sFileName
	 * @return
	 */
	protected String getValidFileName( String _sFileName ) {
		String sNewString = _sFileName.replaceAll("\\:", "");
		sNewString = sNewString.replaceAll("\\[", "(");
		sNewString = sNewString.replaceAll("\\]", ")");
		sNewString = sNewString.replaceAll("\\>", "-");
		return sNewString;
	}

	/**
	 * Opens the export dialog, allowing the user to select the Excel annotation file.
	 * Currently only supports Excel output.
	 * 
	 * @param event
	 */
	public void widgetSelected(SelectionEvent event) 
	{
		FileDialog dlg = new FileDialog(shell,SWT.OPEN);
		String sFileName = getValidFileName(msEntry.getDisplayName()+fileExtension);
		dlg.setFileName(sFileName);
		dlg.setFilterExtensions(new String[] {"*" + fileExtension});
		dlg.setText("File Explorer");

		boolean bDone = false;
		while( ! bDone ) {
			sInputFile = dlg.open();
			try {
				if (sInputFile != null) {
					readPeakAnnotationFile();
					bDone = true;
				} else {
					bDone = true;					
				}
			} catch (NullPointerException e)
			{
				//delete files that were created!
				logger.error(e.getMessage(),e);
				ErrorUtils.createErrorMessageBox(Display.getCurrent().getActiveShell(), "Unable to save file",e);

			} catch (IOException e) {
				//delete files that were created!
				logger.error(e.getMessage(),e);
				ErrorUtils.createErrorMessageBox(Display.getCurrent().getActiveShell(), "Unable to save file",e);
			} catch (JAXBException e) {
				//delete files that were created!
				logger.error(e.getMessage(),e);
				ErrorUtils.createErrorMessageBox(Display.getCurrent().getActiveShell(), "Unable to save file",e);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				logger.error(e.getMessage(),e);
				ErrorUtils.createErrorMessageBox(Display.getCurrent().getActiveShell(), "Unable to save file",e);
			} 
		}
	}

	protected MaorAnnotationInputProcess getNewMaorAnnotationInputProcess() {
		return new MaorAnnotationInputProcess(getInputFile());
	}

	/**
	 * Creates an instance of the annotation reader process, then reads the annotation file.
	 * 
	 * @throws IOException
	 * @throws Exception
	 * @see MaorAnnotationInputProcess
	 */
	protected void readPeakAnnotationFile() throws IOException, Exception {
		//create progress dialog for copying files
		ProgressDialog pDialog = new ProgressDialog(this.shell);
		//fill parameter
		MaorAnnotationInputProcess inputProcess = getNewMaorAnnotationInputProcess();
		//set the worker
		pDialog.setWorker(inputProcess);

		//check Cancel
		if(pDialog.open() == SWT.OK) {
			// if successful, set the label info in the adapter so the Export adapter can get them
			setPeakLabels(inputProcess.getPeakLabels());
			setPeakToLabels(inputProcess.getPeakToLabels());
		}
	}

	/**
	 * @return
	 */
	public Shell getShell() {
		return shell;
	}

	/**
	 * @param shell
	 */
	public void setShell(Shell shell) {
		this.shell = shell;
	}

	/**
	 * @param msEntry
	 */
	public void setMSEntry(Entry msEntry) {
		this.msEntry = msEntry;
	}

	/**
	 * @param fileExtension
	 */
	public void setFileExtension(String fileExtension) {
		this.fileExtension = fileExtension;
	}

	/**
	 * @param exportType
	 */
	public void setExportType(ExportTypes exportType) {
		this.exportType = exportType;
	}

	/**
	 * @return
	 */
	public String getInputFile() {
		return sInputFile;
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
