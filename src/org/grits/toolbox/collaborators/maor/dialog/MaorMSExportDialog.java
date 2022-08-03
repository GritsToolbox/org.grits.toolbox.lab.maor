package org.grits.toolbox.collaborators.maor.dialog;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.List;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.grits.toolbox.collaborators.maor.adaptor.MaorAnnotationFileAdapter;
import org.grits.toolbox.collaborators.maor.adaptor.MaorExportFileAdapter;
import org.grits.toolbox.core.datamodel.Entry;
import org.grits.toolbox.core.datamodel.dialog.ModalDialog;

/**
 * Exports annotated report for Maor (version 0.9)
 * 
 * @author D Brent Weatherly (dbrentw@uga.edu)
 */
public class MaorMSExportDialog extends ModalDialog {

	protected String[] downloadOptions = {"Export Excel file"};
	public enum ExportTypes {Excel};
	
	private Button annotationButton = null;
	private Button okButton;
	private Entry msEntry;
	protected List downloadlist;
	
	private String sSelected;
	private Text txtOutput;
	private MaorExportFileAdapter msExportFileAdapter;
	private MaorAnnotationFileAdapter annotationFileAdapter;

	public MaorMSExportDialog(Shell parentShell) {
		super(parentShell);
		msExportFileAdapter = getNewExportAdapter();
		annotationFileAdapter = getNewMaorAnnotationFileAdapter();
	}
	
	protected MaorExportFileAdapter getNewExportAdapter() {
		MaorExportFileAdapter adapter = new MaorExportFileAdapter();
		return adapter;
	}
	
	protected MaorAnnotationFileAdapter getNewMaorAnnotationFileAdapter() {
		MaorAnnotationFileAdapter adapter = new MaorAnnotationFileAdapter();
		return adapter;
	}
	
	@Override
	public void create()
	{
		super.create();
		setTitle("Export");
		setMessage("Export MS Report");
	}
			
	@Override
	protected Control createDialogArea(final Composite parent) 
	{
		//has to be gridLayout, since it extends TitleAreaDialog
		GridLayout gridLayout = new GridLayout();
		gridLayout.numColumns = 4;
		gridLayout.verticalSpacing = 10;
		parent.setLayout(gridLayout);
		
		this.txtOutput = new Text(parent, SWT.WRAP | SWT.MULTI | SWT.BORDER | SWT.READ_ONLY );
		this.txtOutput.setText("Please select export type.\n");
		this.txtOutput.setFont(boldFont);
		this.txtOutput.setBackground(Display.getCurrent().getSystemColor(SWT.COLOR_WHITE));
		GridData gridDataTxtOutput = new GridData(GridData.FILL_HORIZONTAL);
		gridDataTxtOutput.horizontalSpan = 4;
		gridDataTxtOutput.verticalSpan = 2;
		this.txtOutput.setLayoutData(gridDataTxtOutput);

		/*
		 * First row starts:download list
		 */
		createList(parent);
		createButtonAnnotation(parent);
		createButtonOK(parent);
		
		createButtonCancel(parent);

		return parent;
	}

	protected SelectionListener downloadlistListener = new SelectionListener() {
		@Override
		public void widgetSelected(SelectionEvent e) {
			//enables the download button
//			okButton.setEnabled(true);
			annotationButton.setEnabled(true);
			sSelected = downloadlist.getItem(downloadlist.getSelectionIndex()).toString();
			if(sSelected.equals(downloadOptions[0])) {
				msExportFileAdapter.setFileExtension(".xlsx");
				msExportFileAdapter.setExportType(ExportTypes.Excel);
			}
		}

		@Override
		public void widgetDefaultSelected(SelectionEvent e) {
		}
	};	
	
	protected void createList(Composite parent2) {
		downloadlist = new List(parent2, SWT.SINGLE);
		GridData gridData = new GridData(GridData.FILL_BOTH);
		gridData.horizontalSpan = 4;
		gridData.verticalSpan = 1;
		downloadlist.setLayoutData(gridData);
		//add data to list
		downloadlist.add(downloadOptions[0]);
		downloadlist.select(0);
//		downloadlist.add(downloadOptions[1]);
//		downloadlist.add(downloadOptions[2]);
		//add listener
		downloadlist.addSelectionListener(downloadlistListener);
	}
	
	protected Button createButtonAnnotation(final Composite parent2) {
		//create a gridData for OKButton
		GridData annotData = new GridData(GridData.HORIZONTAL_ALIGN_END);
		annotData.grabExcessHorizontalSpace = true;
		annotData.horizontalSpan = 1;
//		okData.widthHint = 100;
		annotationButton = new Button(parent2, SWT.PUSH);
		annotationButton.setText("Select Annotation File");
		//add export file adaptor
		annotationFileAdapter.setShell(parent2.getShell());
		annotationFileAdapter.setMSEntry(this.msEntry);
		annotationFileAdapter.setFileExtension("xlsx");
		annotationButton.addSelectionListener(annotationFileAdapter);
		annotationButton.addSelectionListener(new SelectionListener() {
			
			@Override
			public void widgetSelected(SelectionEvent e) {
				if( annotationFileAdapter != null && 
						annotationFileAdapter.getPeakLabels() != null &&
						! annotationFileAdapter.getPeakLabels().isEmpty() &&
						annotationFileAdapter.getPeakToLabels() != null &&
						! annotationFileAdapter.getPeakToLabels().isEmpty() ) {
					msExportFileAdapter.setPeakLabels(annotationFileAdapter.getPeakLabels());
					msExportFileAdapter.setPeakToLabels(annotationFileAdapter.getPeakToLabels());
					okButton.setEnabled(true);
				}
				
			}
			
			@Override
			public void widgetDefaultSelected(SelectionEvent e) {
				// TODO Auto-generated method stub
				
			}
		});
		annotationButton.setLayoutData(annotData);
		annotationButton.setEnabled(true);		
		
		return annotationButton;
	}
	
	@Override
	protected Button createButtonOK(final Composite parent2) {
		//create a gridData for OKButton
		GridData okData = new GridData(GridData.HORIZONTAL_ALIGN_END);
		okData.grabExcessHorizontalSpace = true;
		okData.horizontalSpan = 1;
//		okData.widthHint = 100;
		okButton = new Button(parent2, SWT.PUSH);
		okButton.setText("Export");
		//add export file adaptor
		msExportFileAdapter.setShell(parent2.getShell());
		msExportFileAdapter.setMSEntry(this.msEntry);
		msExportFileAdapter.setFileExtension(".xlsx");
		msExportFileAdapter.setExportType(ExportTypes.Excel);
		okButton.addSelectionListener(msExportFileAdapter);
		okButton.setLayoutData(okData);
		okButton.setEnabled(false);		
		
		return okButton;
	}

	@Override
	protected boolean isValidInput() {
		return true;
	}

	@Override
	protected Entry createEntry() {
		return msEntry;
	}

	public void setMSEntry(Entry msEntry) {
		this.msEntry = msEntry;
	}
}
