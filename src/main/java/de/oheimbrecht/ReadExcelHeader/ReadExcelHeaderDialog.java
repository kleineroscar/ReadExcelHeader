/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package de.oheimbrecht.ReadExcelHeader;

import org.eclipse.swt.SWT;
import org.eclipse.swt.custom.CCombo;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.ShellAdapter;
import org.eclipse.swt.events.ShellEvent;
import org.eclipse.swt.events.VerifyEvent;
import org.eclipse.swt.events.VerifyListener;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.pentaho.di.core.Const;
import org.pentaho.di.core.exception.KettleStepException;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.ui.core.dialog.ErrorDialog;
import org.pentaho.di.ui.trans.step.BaseStepDialog;

import org.pentaho.di.trans.step.BaseStepMeta;
import org.pentaho.di.trans.step.StepDialogInterface;

public class ReadExcelHeaderDialog extends BaseStepDialog implements StepDialogInterface {
	// this is the object the stores the step's settings
	// the dialog reads the settings from it when opening
	// the dialog writes the settings to it when confirmed
	private ReadExcelHeaderMeta meta;
	private Label wLabelStepFilename, wLabelStepStartRow, wLabelStepSampleRows;
	private CCombo wComboStepFilename;
	private FormData wFormStepFilename, wFormLabelStepStartRow, wFormStepStartRow, wFormLabelStepSampleRows, wFormStepSampleRows;
	private Text wTextStartRow, wTextSampleRows;
	RowMetaInterface inputSteps;

	public ReadExcelHeaderDialog(Shell parent, Object in, TransMeta transMeta, String sname) {
		super(parent, (BaseStepMeta) in, transMeta, sname);
		meta = (ReadExcelHeaderMeta) in;
	}

	/**
	 * This method is called by Spoon when the user opens the settings dialog of the
	 * step. It should open the dialog and return only once the dialog has been
	 * closed by the user.
	 * 
	 * If the user confirms the dialog, the meta object (passed in the constructor)
	 * must be updated to reflect the new step settings. The changed flag of the
	 * meta object must reflect whether the step configuration was changed by the
	 * dialog.
	 * 
	 * If the user cancels the dialog, the meta object must not be updated, and its
	 * changed flag must remain unaltered.
	 * 
	 * The open() method must return the name of the step after the user has
	 * confirmed the dialog, or null if the user cancelled the dialog.
	 */
	public String open() {
		// store some convenient SWT variables
		Shell parent = getParent();
		Display display = parent.getDisplay();

		// SWT code for preparing the dialog
		shell = new Shell(parent, SWT.DIALOG_TRIM | SWT.RESIZE | SWT.MIN | SWT.MAX);
		props.setLook(shell);
		setShellImage(shell, meta);

		// Save the value of the changed flag on the meta object. If the user cancels
		// the dialog, it will be restored to this saved value.
		// The "changed" variable is inherited from BaseStepDialog
		changed = meta.hasChanged();

		// The ModifyListener used on all controls. It will update the meta object to
		// indicate that changes are being made.
		ModifyListener lsMod = new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				meta.setChanged();
			}
		};

		// ------------------------------------------------------- //
		// SWT code for building the actual settings dialog //
		// ------------------------------------------------------- //
		FormLayout formLayout = new FormLayout();
		formLayout.marginWidth = Const.FORM_MARGIN;
		formLayout.marginHeight = Const.FORM_MARGIN;
		shell.setLayout(formLayout);
		shell.setText(Messages.getString("ReadExcelHeaderDialog.Shell.Title"));
		int middle = props.getMiddlePct();
		int margin = Const.MARGIN;

		// Stepname line
		wlStepname = new Label(shell, SWT.RIGHT);
		wlStepname.setText(Messages.getString("ReadExcelHeaderDialog.StepName.Label"));
		props.setLook(wlStepname);
		fdlStepname = new FormData();
		fdlStepname.left = new FormAttachment(0, 0);
		fdlStepname.right = new FormAttachment(middle, -margin);
		fdlStepname.top = new FormAttachment(0, margin);
		wlStepname.setLayoutData(fdlStepname);

		wStepname = new Text(shell, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		wStepname.setText(stepname);
		props.setLook(wStepname);
		wStepname.addModifyListener(lsMod);
		fdStepname = new FormData();
		fdStepname.left = new FormAttachment(middle, 0);
		fdStepname.top = new FormAttachment(0, margin);
		fdStepname.right = new FormAttachment(100, 0);
		wStepname.setLayoutData(fdStepname);

		// OK and cancel buttons
		wOK = new Button(shell, SWT.PUSH);
		wOK.setText(Messages.getString("System.Button.OK"));
		wCancel = new Button(shell, SWT.PUSH);
		wCancel.setText(Messages.getString("System.Button.Cancel")); //$NON-NLS-1$
		
		// Add listeners
		lsCancel   = new Listener() { public void handleEvent(Event e) { cancel(); } };
		lsOK       = new Listener() { public void handleEvent(Event e) { ok();     } };

		// Filename Input Step
		wLabelStepFilename = new Label(shell, SWT.RIGHT);
		wLabelStepFilename.setText(Messages.getString("ReadExcelHedaerDialog.FilenameField.Label"));
		props.setLook(wLabelStepFilename);
		wFormStepFilename = new FormData();
		wFormStepFilename.left = new FormAttachment(0, 0);
		wFormStepFilename.right = new FormAttachment(middle, -margin);
		wFormStepFilename.top = new FormAttachment(wStepname, margin);
		wLabelStepFilename.setLayoutData(wFormStepFilename);
		wComboStepFilename = new CCombo(shell, SWT.BORDER | SWT.READ_ONLY);
		try {
			inputSteps = transMeta.getPrevStepFields(stepMeta);
			wComboStepFilename.setItems(inputSteps.getFieldNames());
		} catch (KettleStepException e1) {
			e1.printStackTrace();
			new ErrorDialog(shell, Messages.getString("ReadExcelHeader.Step.Name"), e1.getStackTrace().toString(), e1);
		}

		props.setLook(wComboStepFilename);
		wComboStepFilename.addModifyListener(lsMod);
		wFormStepFilename = new FormData();
		wFormStepFilename.left = new FormAttachment(middle, 0);
		wFormStepFilename.right = new FormAttachment(100, 0);
		wFormStepFilename.top = new FormAttachment(wStepname, margin);
		wComboStepFilename.setLayoutData(wFormStepFilename);
		
		// start row line
		wLabelStepStartRow = new Label(shell, SWT.RIGHT);
		wLabelStepStartRow.setText(Messages.getString("ReadExcelHeaderDialog.StartRow.Label"));
		props.setLook(wLabelStepStartRow);
		wFormLabelStepStartRow = new FormData();
		wFormLabelStepStartRow.left = new FormAttachment(0, 0);
		wFormLabelStepStartRow.right = new FormAttachment(middle, -margin);
		wFormLabelStepStartRow.top = new FormAttachment(wComboStepFilename, margin);
		wLabelStepStartRow.setLayoutData(wFormLabelStepStartRow);
	
		wTextStartRow = new Text(shell, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wStepname);
//		wTextStartRow.addVerifyListener(new VerifyListener() {
//
//			@Override
//			public void verifyText(VerifyEvent e) {
//				 Text text = (Text)e.getSource();
//
//		         // get old text and create new text by using the VerifyEvent.text
//		         final String oldS = text.getText();
//		         if (oldS.length()==0) {
//		        	 return;
//		         }
//		         String newS = oldS.substring(0, e.start) + e.text + oldS.substring(e.end);
//
//		         boolean isValid = true;
//		         try
//		         {
//		             if (Integer.parseInt(newS) < 0) {
//		            	 throw new NumberFormatException();
//		             }
//		         }
//		         catch(NumberFormatException ex)
//		         {
//		        	 isValid = false;
//		         }
//
////		         System.out.println(newS);
//
//		         if(!isValid)
//		             e.doit = false;
//				
//			}
//			
//		});
		props.setLook(wTextStartRow);
		wTextStartRow.addModifyListener(lsMod);
		wFormStepStartRow = new FormData();
		wFormStepStartRow.left = new FormAttachment(middle, 0);
		wFormStepStartRow.top = new FormAttachment(wComboStepFilename, margin);
		wFormStepStartRow.right = new FormAttachment(100, 0);
		wTextStartRow.setLayoutData(wFormStepStartRow);

		// sample rows line
		wLabelStepSampleRows = new Label(shell, SWT.RIGHT);
		wLabelStepSampleRows.setText(Messages.getString("ReadExcelHeaderDialog.SampleRows.Label"));
		props.setLook(wLabelStepSampleRows);
		wFormLabelStepSampleRows = new FormData();
		wFormLabelStepSampleRows.left = new FormAttachment(0, 0);
		wFormLabelStepSampleRows.right = new FormAttachment(middle, -margin);
		wFormLabelStepSampleRows.top = new FormAttachment(wTextStartRow, margin);
		wLabelStepSampleRows.setLayoutData(wFormLabelStepSampleRows);
		
		wTextSampleRows = new Text(shell, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wStepname);
		props.setLook(wTextSampleRows);
		wTextSampleRows.addModifyListener(lsMod);
		wFormStepSampleRows = new FormData();
		wFormStepSampleRows.left = new FormAttachment(middle, 0);
		wFormStepSampleRows.top = new FormAttachment(wTextStartRow, margin);
		wFormStepSampleRows.right = new FormAttachment(100, 0);
		wTextSampleRows.setLayoutData(wFormStepSampleRows);
		
		BaseStepDialog.positionBottomButtons(shell, new Button[] { wOK, wCancel}, margin, wTextStartRow);


		// Add listeners for cancel and OK
		lsCancel = new Listener() {
			public void handleEvent(Event e) {
				cancel();
			}
		};
		lsOK = new Listener() {
			public void handleEvent(Event e) {
				ok();
			}
		};
		wCancel.addListener(SWT.Selection, lsCancel);
		wOK.addListener(SWT.Selection, lsOK);

		// default listener (for hitting "enter")
		lsDef = new SelectionAdapter() {
			public void widgetDefaultSelected(SelectionEvent e) {
				ok();
			}
		};
		wStepname.addSelectionListener(lsDef);
		wComboStepFilename.addSelectionListener(lsDef);
		wTextStartRow.addSelectionListener(lsDef);

		// Detect X or ALT-F4 or something that kills this window and cancel the dialog
		// properly
		shell.addShellListener(new ShellAdapter() {
			public void shellClosed(ShellEvent e) {
				cancel();
			}
		});

		// Set/Restore the dialog size based on last position on screen
		// The setSize() method is inherited from BaseStepDialog
		setSize();

		// restore the changed flag to original value, as the modify listeners fire
		// during dialog population
		getData();

		// open dialog and enter event loop
		shell.open();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}

		// at this point the dialog has closed, so either ok() or cancel() have been
		// executed
		// The "stepname" variable is inherited from BaseStepDialog
		return stepname;
	}
	
	// Read data from input
	public void getData()
	{
		wTextStartRow.setText(meta.getStartRow());
		wTextSampleRows.setText(meta.getSampleRows());
		try {
			wComboStepFilename.select(Integer.parseInt(meta.getFilenameField()));
		} catch (Exception e) {
		}
		}

	private void cancel() {
		// The "stepname" variable will be the return value for the open() method.
		// Setting to null to indicate that dialog was cancelled.
		stepname = null;
		// Restoring original "changed" flag on the met aobject
		meta.setChanged(changed);
		// close the SWT dialog window
		dispose();
	}

	/**
	 * Called when the user confirms the dialog
	 */
	private void ok() {
		// The "stepname" variable will be the return value for the open() method.
		// Setting to step name from the dialog control
		stepname = wStepname.getText();
		meta.setFilenameField(String.valueOf(wComboStepFilename.getSelectionIndex()));

		String startRowText = wTextStartRow.getText();
		meta.setStartRow((startRowText.length() > 0) ? startRowText : "0");

		String sampleRowsText = wTextSampleRows.getText();
		meta.setSampleRows((sampleRowsText.length() > 0) ? sampleRowsText : "1");
		
		meta.setChanged();
		// close the SWT dialog window
		dispose();
	}
}
