package org.grits.toolbox.collaborators.maor.handler;

import javax.inject.Named;

import org.apache.log4j.Logger;
import org.eclipse.e4.core.di.annotations.CanExecute;
import org.eclipse.e4.core.di.annotations.Execute;
import org.eclipse.e4.ui.model.application.ui.basic.MPart;
import org.eclipse.e4.ui.services.IServiceConstants;
import org.eclipse.e4.ui.workbench.modeling.EPartService;
import org.eclipse.jface.window.Window;
import org.eclipse.swt.widgets.Shell;
import org.grits.toolbox.collaborators.maor.dialog.MaorMSExportDialog;
import org.grits.toolbox.core.dataShare.PropertyHandler;
import org.grits.toolbox.core.datamodel.Entry;
import org.grits.toolbox.core.utilShare.ErrorUtils;
import org.grits.toolbox.entry.ms.views.tabbed.MassSpecMultiPageViewer;

/**
 * Export command for Maor's work. Call MSExportDialog
 * 
 * @author dbrentw
 * 
 */
public class MaorMSExportCommand {
	private static final Logger logger = Logger.getLogger(MaorMSExportCommand.class);

	private Entry entry = null;

	@Execute
	public void execute(@Named(IServiceConstants.ACTIVE_PART) MPart part,
			@Named (IServiceConstants.ACTIVE_SHELL) Shell shell,
			EPartService partService) {

		if (initialize(part, partService)) {
			createExportDialog(shell);
		} else {
			logger.warn("A valid MS Annotation entry must be open and active in order to export.");
			ErrorUtils.createWarningMessageBox(
					shell, "Invalid Entry",	"An appropriate MS Annotation entry must be open and active in order to export.");
		}
	}

	/**
	 * Create instance of the export dialog and open it.
	 * 
	 * @param activeShell
	 * @see MaorMSExportDialog
	 */
	protected void createExportDialog(Shell activeShell) {
		MaorMSExportDialog dialog = getNewExportDialog (activeShell);
		// set parent entry
		dialog.setMSEntry(entry);
		if (dialog.open() == Window.OK) {
			// to do something..
		}
	}
	
	/**
	 * @param activeShell
	 * @return
	 */
	protected MaorMSExportDialog getNewExportDialog (Shell activeShell) {
		return new MaorMSExportDialog(PropertyHandler.getModalDialog(activeShell));
	}
	
	/**
	 * @param part
	 * @param partService
	 * @return
	 */
	protected boolean initialize(MPart part, EPartService partService) {
		try {
			MassSpecMultiPageViewer viewer = null;
			if (part != null && part.getObject() instanceof MassSpecMultiPageViewer) {
				viewer = (MassSpecMultiPageViewer) part.getObject();
			} else { // try to find an open part of the required type
				for (MPart mPart: partService.getParts()) {
					if (mPart.getObject() instanceof MassSpecMultiPageViewer) {
						if (mPart.equals(mPart.getParent().getSelectedElement())) {
							viewer = (MassSpecMultiPageViewer) mPart.getObject();
							if (!viewer.getPeaksView().isEmpty())
								break;
						}
					}
				}
			}
			if (viewer != null) {
				if( viewer.getScansView() == null || viewer.getScansView().getViewBase() == null) {
					return false;
				}
				setEntry(viewer.getEntry());

				return true;
			} else {
				return false;
			}
		} catch( Exception e ) {
			logger.error(e.getMessage(), e);
			return false;
		}		
	}

	/**
	 * @return
	 */
	public Entry getEntry() {
		return entry;
	}

	/**
	 * @param entry
	 */
	public void setEntry(Entry entry) {
		this.entry = entry;
	}

	@CanExecute
	public boolean isEnabled(@Named(IServiceConstants.ACTIVE_PART) MPart part, EPartService partService) {
		return initialize(part, partService);
	}
}
