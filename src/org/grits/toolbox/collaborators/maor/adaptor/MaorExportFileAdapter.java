package org.grits.toolbox.collaborators.maor.adaptor;

import java.io.File;
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
import org.grits.toolbox.collaborators.maor.process.MaorExportProcess;
import org.grits.toolbox.core.datamodel.Entry;
import org.grits.toolbox.core.utilShare.ErrorUtils;
import org.grits.toolbox.entry.ms.property.MassSpecEntityProperty;
import org.grits.toolbox.entry.ms.property.MassSpecProperty;
import org.grits.toolbox.widgets.processDialog.ProgressDialog;

/**
 * Maor Excel report file adapter. Currently only supports exporting into Excel format.
 * 
 * @author D Brent Weatherly (dbrentw@uga.edu)
 *
 */
public class MaorExportFileAdapter extends SelectionAdapter {

	//log4J Logger
	private static final Logger logger = Logger.getLogger(MaorExportFileAdapter.class);

	protected Shell shell = null;
	protected String fileExtension = null;
	protected Entry msEntry = null;
	protected String sOutputFile = null;
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
	 * Opens the export dialog, allowing the user to select the Excel output file and then creates the report.
	 * Currently only supports Excel output.
	 * 
	 * @param event
	 */
	@Override
	public void widgetSelected(SelectionEvent event)  {
		FileDialog dlg = new FileDialog(shell,SWT.SAVE);
		String sFileName = getValidFileName(msEntry.getDisplayName()+fileExtension);
		dlg.setFileName(sFileName);
		dlg.setFilterExtensions(new String[] {"*" + fileExtension});
		dlg.setText("File Explorer");

		boolean bDone = false;
		while( ! bDone ) {
			sOutputFile = dlg.open();
			try {
				if (sOutputFile != null) {
					int iRes = SWT.OK;
					File f = new File(sOutputFile);
					if( f.exists() ) {
						String sEMsg = "The selected export file exists.";
						iRes = ErrorUtils.createMultiConfirmationMessageBoxReturn(
								Display.getCurrent().getActiveShell(), 
								sEMsg, "Overwrite?", false);

					}

					if( iRes == SWT.OK ) {
						if ( exportType == ExportTypes.Excel ) {
							exportExcel();
						} else {
							
						}
					}
					//close
					if( ! (iRes == SWT.NO) ) {
						bDone = true;
						shell.close();
					}
				} else {
					bDone = true;					
				}
			} catch (NullPointerException e) {
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

	/**
	 * Creates a new instance of the export processor.
	 * 
	 * @return
	 */
	protected MaorExportProcess getNewExportProcess() {
		return new MaorExportProcess( getPeakLabels(), getPeakToLabels() );
	}

	/**
	 * Method to export an Excel version of the Maor report
	 * 
	 * @throws IOException
	 * @throws Exception
	 */
	protected void exportExcel() throws IOException, Exception {
		//create progress dialog for copying files
		ProgressDialog t_dialog = new ProgressDialog(this.shell);
		//fill parameter
		MaorExportProcess t_worker = getNewExportProcess();
		t_worker.setOutputFile(sOutputFile);

		MassSpecEntityProperty entityProp = (MassSpecEntityProperty) this.msEntry.getProperty();
		String sMzXMLFile = entityProp.getMassSpecParentProperty().getFullyQualifiedFolderName(msEntry) + File.separator + entityProp.getDataFile().getName();
//		String sInputFile = ((MassSpecProperty) msEntry.getProperty().getParentProperty()).getFullyQualifiedMzXMLFileName(msEntry);
		t_worker.setInputFile(sMzXMLFile);
		//set the worker
		t_dialog.setWorker(t_worker);

		//check Cancel
		if(t_dialog.open() != SWT.OK)
		{
			//delete the file
			// DBW: I commented this out, allowing for creation of a partial report if the user hits Cancel
//			new File(sOutputFile).delete();
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
	 * @return
	 */
	public String getOutputFile() {
		return sOutputFile;
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
