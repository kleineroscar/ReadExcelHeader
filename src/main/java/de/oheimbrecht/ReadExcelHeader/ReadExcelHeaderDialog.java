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
import org.eclipse.swt.custom.CTabFolder;
import org.eclipse.swt.custom.CTabItem;
import org.eclipse.swt.events.FocusListener;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.ShellAdapter;
import org.eclipse.swt.events.ShellEvent;
import org.eclipse.swt.graphics.Cursor;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.pentaho.di.core.Const;
import org.pentaho.di.core.Props;
import org.pentaho.di.core.exception.KettleException;
// import org.pentaho.di.core.exception.KettleStepException;
import org.pentaho.di.core.fileinput.FileInputList;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.util.Utils;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.ui.core.widget.ColumnInfo;
import org.pentaho.di.ui.core.widget.TableView;
import org.pentaho.di.ui.core.widget.TextVar;
import org.pentaho.di.ui.core.dialog.EnterSelectionDialog;
import org.pentaho.di.ui.core.dialog.ErrorDialog;
import org.pentaho.di.ui.core.events.dialog.FilterType;
import org.pentaho.di.ui.core.events.dialog.SelectionAdapterFileDialogTextVar;
import org.pentaho.di.ui.core.events.dialog.SelectionAdapterOptions;
import org.pentaho.di.ui.core.events.dialog.SelectionOperation;
import org.pentaho.di.ui.trans.step.BaseStepDialog;

import org.pentaho.di.trans.step.BaseStepMeta;
import org.pentaho.di.trans.step.StepDialogInterface;

public class ReadExcelHeaderDialog extends BaseStepDialog implements StepDialogInterface {
	// this is the object the stores the step's settings
	// the dialog reads the settings from it when opening
	// the dialog writes the settings to it when confirmed
	private CTabFolder wTabFolder;
	private FormData fdTabFolder;

	private CTabItem wFileTab, wContentTab;

	private Composite wFileComp, wContentComp;
	private FormData fdFileComp, fdContentComp;

	private Label wlExcludeFilemask;
	private TextVar wExcludeFilemask;
	private FormData fdlExcludeFilemask, fdExcludeFilemask;

	private Label wlFilename;
	private Button wbbFilename; // Browse: add file or directory

	private Button wbdFilename; // Delete
	private Button wbeFilename; // Edit
	private Button wbaFilename; // Add or change
	private TextVar wFilename;
	private FormData fdlFilename, fdbFilename, fdbdFilename, fdbeFilename, fdbaFilename, fdFilename;

	private Label wlFilenameList;
	private TableView wFilenameList;
	private FormData fdlFilenameList, fdFilenameList;

	private Label wlFilemask;
	private TextVar wFilemask;
	private FormData fdlFilemask, fdFilemask;

	private Button wbShowFiles;
	private FormData fdbShowFiles;

	private Group wOriginFiles, wStartRowGroup, wSampleRowsGroup;

	private FormData fdOriginFiles, fdFilenameField, fdlFilenameField;
	private Button wFileField;

	private Label wlFileField, wlFilenameField;
	private CCombo wFilenameField;
	private FormData fdlFileField, fdFileField;

	private Label wlWildcardField;
	private CCombo wWildcardField;
	private FormData fdlWildcardField, fdWildcardField;

	private Label wlExcludeWildcardField;
	private CCombo wExcludeWildcardField;
	private FormData fdlExcludeWildcardField, fdExcludeWildcardField;

	private Label wlIncludeSubFolder;
	private FormData fdlIncludeSubFolder;
	private Button wIncludeSubFolder;
	private FormData fdIncludeSubFolder;

	private Label wldoNotFailIfNoFile;
	private Button wdoNotFailIfNoFile;
	private FormData fdldoNotFailIfNoFile, fddoNotFailIfNoFile;

	private ReadExcelHeaderMeta meta;
	private Label wLabelStepStartRow, wLabelStepSampleRows;
	private FormData wFormLabelStepStartRow, wFormStepStartRow, wFormLabelStepSampleRows, wFormStepSampleRows,
			fdStartRow, fdSampleRows;
	private TextVar wTextStartRow, wTextSampleRows;
	RowMetaInterface inputSteps;
	private Label wlStartRowField;
	private FormData fdlStartRowField;
	private Button wStartRowField;
	private FormData fdStartRowField;
	private Label wlStartRowSelField;
	private FormData fdlStartRowSelField;
	private CCombo wStartRowSelField;
	private FormData fdStartRowSelField;
	private Label wSeparator;
	private FormData fdSeparator;

	private boolean getpreviousFields = false;

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

		wTabFolder = new CTabFolder(shell, SWT.BORDER);
		props.setLook(wTabFolder, Props.WIDGET_STYLE_TAB);

		// ////////////////////////
		// START OF FILE TAB ///
		// ////////////////////////
		wFileTab = new CTabItem(wTabFolder, SWT.NONE);
		wFileTab.setText(Messages.getString("ReadExcelHeaderDialog.File.Tab"));

		wFileComp = new Composite(wTabFolder, SWT.NONE);
		props.setLook(wFileComp);

		FormLayout fileLayout = new FormLayout();
		fileLayout.marginWidth = 3;
		fileLayout.marginHeight = 3;
		wFileComp.setLayout(fileLayout);

		// ///////////////////////////////
		// START OF Origin files GROUP //
		// ///////////////////////////////

		wOriginFiles = new Group(wFileComp, SWT.SHADOW_NONE);
		props.setLook(wOriginFiles);
		wOriginFiles.setText(Messages.getString("ReadExcelHeaderDialog.wOriginFiles.Label"));

		FormLayout OriginFilesgroupLayout = new FormLayout();
		OriginFilesgroupLayout.marginWidth = 10;
		OriginFilesgroupLayout.marginHeight = 10;
		wOriginFiles.setLayout(OriginFilesgroupLayout);

		// Is Filename defined in a Field
		wlFileField = new Label(wOriginFiles, SWT.RIGHT);
		wlFileField.setText(Messages.getString("ReadExcelHeaderDialog.FileField.Label"));
		props.setLook(wlFileField);
		fdlFileField = new FormData();
		fdlFileField.left = new FormAttachment(0, -margin);
		fdlFileField.top = new FormAttachment(0, margin);
		fdlFileField.right = new FormAttachment(middle, -2 * margin);
		wlFileField.setLayoutData(fdlFileField);

		wFileField = new Button(wOriginFiles, SWT.CHECK);
		props.setLook(wFileField);
		wFileField.setToolTipText(Messages.getString("ReadExcelHeaderDialog.FileField.Tooltip"));
		fdFileField = new FormData();
		fdFileField.left = new FormAttachment(middle, -margin);
		fdFileField.top = new FormAttachment(0, margin);
		wFileField.setLayoutData(fdFileField);
		SelectionAdapter lfilefield = new SelectionAdapter() {
			public void widgetSelected(SelectionEvent arg0) {
				ActiveFileField();
				setFileField();
				meta.setChanged(true);
			}
		};
		wFileField.addSelectionListener(lfilefield);

		// Filename field
		wlFilenameField = new Label(wOriginFiles, SWT.RIGHT);
		wlFilenameField.setText(Messages.getString("ReadExcelHeaderDialog.FilenameField.Label"));
		props.setLook(wlFilenameField);
		fdlFilenameField = new FormData();
		fdlFilenameField.left = new FormAttachment(0, -margin);
		fdlFilenameField.top = new FormAttachment(wFileField, margin);
		fdlFilenameField.right = new FormAttachment(middle, -2 * margin);
		wlFilenameField.setLayoutData(fdlFilenameField);

		wFilenameField = new CCombo(wOriginFiles, SWT.BORDER | SWT.READ_ONLY);
		wFilenameField.setEditable(true);
		props.setLook(wFilenameField);
		wFilenameField.addModifyListener(lsMod);
		fdFilenameField = new FormData();
		fdFilenameField.left = new FormAttachment(middle, -margin);
		fdFilenameField.top = new FormAttachment(wFileField, margin);
		fdFilenameField.right = new FormAttachment(100, -margin);
		wFilenameField.setLayoutData(fdFilenameField);
		wFilenameField.addFocusListener(new FocusListener() {
			public void focusLost(org.eclipse.swt.events.FocusEvent e) {
			}

			public void focusGained(org.eclipse.swt.events.FocusEvent e) {
				Cursor busy = new Cursor(shell.getDisplay(), SWT.CURSOR_WAIT);
				shell.setCursor(busy);
				setFileField();
				shell.setCursor(null);
				busy.dispose();
			}
		});

		// Wildcard field
		wlWildcardField = new Label(wOriginFiles, SWT.RIGHT);
		wlWildcardField.setText(Messages.getString("ReadExcelHeaderDialog.wlWildcardField.Label"));
		props.setLook(wlWildcardField);
		fdlWildcardField = new FormData();
		fdlWildcardField.left = new FormAttachment(0, -margin);
		fdlWildcardField.top = new FormAttachment(wFilenameField, margin);
		fdlWildcardField.right = new FormAttachment(middle, -2 * margin);
		wlWildcardField.setLayoutData(fdlWildcardField);

		wWildcardField = new CCombo(wOriginFiles, SWT.BORDER | SWT.READ_ONLY);
		wWildcardField.setEditable(true);
		props.setLook(wWildcardField);
		wWildcardField.addModifyListener(lsMod);
		fdWildcardField = new FormData();
		fdWildcardField.left = new FormAttachment(middle, -margin);
		fdWildcardField.top = new FormAttachment(wFilenameField, margin);
		fdWildcardField.right = new FormAttachment(100, -margin);
		wWildcardField.setLayoutData(fdWildcardField);

		// ExcludeWildcard field
		wlExcludeWildcardField = new Label(wOriginFiles, SWT.RIGHT);
		wlExcludeWildcardField.setText(Messages.getString("ReadExcelHeaderDialog.wlExcludeWildcardField.Label"));
		props.setLook(wlExcludeWildcardField);
		fdlExcludeWildcardField = new FormData();
		fdlExcludeWildcardField.left = new FormAttachment(0, -margin);
		fdlExcludeWildcardField.top = new FormAttachment(wWildcardField, margin);
		fdlExcludeWildcardField.right = new FormAttachment(middle, -2 * margin);
		wlExcludeWildcardField.setLayoutData(fdlExcludeWildcardField);

		wExcludeWildcardField = new CCombo(wOriginFiles, SWT.BORDER | SWT.READ_ONLY);
		wExcludeWildcardField.setEditable(true);
		props.setLook(wExcludeWildcardField);
		wExcludeWildcardField.addModifyListener(lsMod);
		fdExcludeWildcardField = new FormData();
		fdExcludeWildcardField.left = new FormAttachment(middle, -margin);
		fdExcludeWildcardField.top = new FormAttachment(wWildcardField, margin);
		fdExcludeWildcardField.right = new FormAttachment(100, -margin);
		wExcludeWildcardField.setLayoutData(fdExcludeWildcardField);

		// Is includeSubFoldername defined in a Field
		wlIncludeSubFolder = new Label(wOriginFiles, SWT.RIGHT);
		wlIncludeSubFolder.setText(Messages.getString("ReadExcelHeaderDialog.includeSubFolder.Label"));
		props.setLook(wlIncludeSubFolder);
		fdlIncludeSubFolder = new FormData();
		fdlIncludeSubFolder.left = new FormAttachment(0, -margin);
		fdlIncludeSubFolder.top = new FormAttachment(wExcludeWildcardField, margin);
		fdlIncludeSubFolder.right = new FormAttachment(middle, -2 * margin);
		wlIncludeSubFolder.setLayoutData(fdlIncludeSubFolder);

		wIncludeSubFolder = new Button(wOriginFiles, SWT.CHECK);
		props.setLook(wIncludeSubFolder);
		wIncludeSubFolder.setToolTipText(Messages.getString("ReadExcelHeaderDialog.includeSubFolder.Tooltip"));
		fdIncludeSubFolder = new FormData();
		fdIncludeSubFolder.left = new FormAttachment(middle, -margin);
		fdIncludeSubFolder.top = new FormAttachment(wExcludeWildcardField, margin);
		wIncludeSubFolder.setLayoutData(fdIncludeSubFolder);
		wIncludeSubFolder.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent selectionEvent) {
				meta.setChanged();
			}
		});

		fdOriginFiles = new FormData();
		fdOriginFiles.left = new FormAttachment(0, margin);
		fdOriginFiles.top = new FormAttachment(wFilenameList, margin);
		fdOriginFiles.right = new FormAttachment(100, -margin);
		wOriginFiles.setLayoutData(fdOriginFiles);

		// ///////////////////////////////////////////////////////////
		// / END OF Origin files GROUP
		// ///////////////////////////////////////////////////////////

		// Filename line
		wlFilename = new Label(wFileComp, SWT.RIGHT);
		wlFilename.setText(Messages.getString("ReadExcelHeaderDialog.Filename.Label"));
		props.setLook(wlFilename);
		fdlFilename = new FormData();
		fdlFilename.left = new FormAttachment(0, 0);
		fdlFilename.top = new FormAttachment(wOriginFiles, margin);
		fdlFilename.right = new FormAttachment(middle, -margin);
		wlFilename.setLayoutData(fdlFilename);

		wbbFilename = new Button(wFileComp, SWT.PUSH | SWT.CENTER);
		props.setLook(wbbFilename);
		wbbFilename.setText(Messages.getString("ReadExcelHeaderDialog.FilenameBrowse.Button"));
		wbbFilename.setToolTipText(Messages.getString("System.Tooltip.BrowseForFileOrDirAndAdd"));
		fdbFilename = new FormData();
		fdbFilename.right = new FormAttachment(100, 0);
		fdbFilename.top = new FormAttachment(wOriginFiles, margin);
		wbbFilename.setLayoutData(fdbFilename);

		wbaFilename = new Button(wFileComp, SWT.PUSH | SWT.CENTER);
		props.setLook(wbaFilename);
		wbaFilename.setText(Messages.getString("ReadExcelHeaderDialog.FilenameAdd.Button"));
		wbaFilename.setToolTipText(Messages.getString("ReadExcelHeaderDialog.FilenameAdd.Tooltip"));
		fdbaFilename = new FormData();
		fdbaFilename.right = new FormAttachment(wbbFilename, -margin);
		fdbaFilename.top = new FormAttachment(wOriginFiles, margin);
		wbaFilename.setLayoutData(fdbaFilename);

		wFilename = new TextVar(transMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wFilename);
		wFilename.addModifyListener(lsMod);
		fdFilename = new FormData();
		fdFilename.left = new FormAttachment(middle, 0);
		fdFilename.right = new FormAttachment(wbaFilename, -margin);
		fdFilename.top = new FormAttachment(wOriginFiles, margin);
		wFilename.setLayoutData(fdFilename);

		wlFilemask = new Label(wFileComp, SWT.RIGHT);
		wlFilemask.setText(Messages.getString("ReadExcelHeaderDialog.RegExp.Label"));
		props.setLook(wlFilemask);
		fdlFilemask = new FormData();
		fdlFilemask.left = new FormAttachment(0, 0);
		fdlFilemask.top = new FormAttachment(wFilename, margin);
		fdlFilemask.right = new FormAttachment(middle, -margin);
		wlFilemask.setLayoutData(fdlFilemask);
		wFilemask = new TextVar(transMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wFilemask);
		wFilemask.addModifyListener(lsMod);
		fdFilemask = new FormData();
		fdFilemask.left = new FormAttachment(middle, 0);
		fdFilemask.top = new FormAttachment(wFilename, margin);
		fdFilemask.right = new FormAttachment(100, 0);
		wFilemask.setLayoutData(fdFilemask);

		wlExcludeFilemask = new Label(wFileComp, SWT.RIGHT);
		wlExcludeFilemask.setText(Messages.getString("ReadExcelHeaderDialog.ExcludeFilemask.Label"));
		props.setLook(wlExcludeFilemask);
		fdlExcludeFilemask = new FormData();
		fdlExcludeFilemask.left = new FormAttachment(0, 0);
		fdlExcludeFilemask.top = new FormAttachment(wFilemask, margin);
		fdlExcludeFilemask.right = new FormAttachment(middle, -margin);
		wlExcludeFilemask.setLayoutData(fdlExcludeFilemask);
		wExcludeFilemask = new TextVar(transMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wExcludeFilemask);
		wExcludeFilemask.addModifyListener(lsMod);
		fdExcludeFilemask = new FormData();
		fdExcludeFilemask.left = new FormAttachment(middle, 0);
		fdExcludeFilemask.top = new FormAttachment(wFilemask, margin);
		fdExcludeFilemask.right = new FormAttachment(wFilename, 0, SWT.RIGHT);
		wExcludeFilemask.setLayoutData(fdExcludeFilemask);

		// Filename list line
		wlFilenameList = new Label(wFileComp, SWT.RIGHT);
		wlFilenameList.setText(Messages.getString("ReadExcelHeaderDialog.FilenameList.Label"));
		props.setLook(wlFilenameList);
		fdlFilenameList = new FormData();
		fdlFilenameList.left = new FormAttachment(0, 0);
		fdlFilenameList.top = new FormAttachment(wExcludeFilemask, margin);
		fdlFilenameList.right = new FormAttachment(middle, -margin);
		wlFilenameList.setLayoutData(fdlFilenameList);

		// Buttons to the right of the screen...
		wbdFilename = new Button(wFileComp, SWT.PUSH | SWT.CENTER);
		props.setLook(wbdFilename);
		wbdFilename.setText(Messages.getString("ReadExcelHeaderDialog.FilenameRemove.Button"));
		wbdFilename.setToolTipText(Messages.getString("ReadExcelHeaderDialog.FilenameRemove.Tooltip"));
		fdbdFilename = new FormData();
		fdbdFilename.right = new FormAttachment(100, 0);
		fdbdFilename.top = new FormAttachment(wExcludeFilemask, 40);
		wbdFilename.setLayoutData(fdbdFilename);

		wbeFilename = new Button(wFileComp, SWT.PUSH | SWT.CENTER);
		props.setLook(wbeFilename);
		wbeFilename.setText(Messages.getString("ReadExcelHeaderDialog.FilenameEdit.Button"));
		wbeFilename.setToolTipText(Messages.getString("ReadExcelHeaderDialog.FilenameEdit.Tooltip"));
		fdbeFilename = new FormData();
		fdbeFilename.right = new FormAttachment(100, 0);
		fdbeFilename.top = new FormAttachment(wbdFilename, margin);
		wbeFilename.setLayoutData(fdbeFilename);

		wbShowFiles = new Button(wFileComp, SWT.PUSH | SWT.CENTER);
		props.setLook(wbShowFiles);
		wbShowFiles.setText(Messages.getString("ReadExcelHeaderDialog.ShowFiles.Button"));
		fdbShowFiles = new FormData();
		fdbShowFiles.left = new FormAttachment(middle, 0);
		fdbShowFiles.bottom = new FormAttachment(100, 0);
		wbShowFiles.setLayoutData(fdbShowFiles);

		ColumnInfo[] colinfo = new ColumnInfo[5];
		colinfo[0] = new ColumnInfo(Messages.getString("ReadExcelHeaderDialog.Files.Filename.Column"),
				ColumnInfo.COLUMN_TYPE_TEXT, false);
		colinfo[1] = new ColumnInfo(Messages.getString("ReadExcelHeaderDialog.Files.Wildcard.Column"),
				ColumnInfo.COLUMN_TYPE_TEXT, false);
		colinfo[2] = new ColumnInfo(Messages.getString("ReadExcelHeaderDialog.Files.ExcludeWildcard.Column"),
				ColumnInfo.COLUMN_TYPE_TEXT, false);

		colinfo[3] = new ColumnInfo(Messages.getString("ReadExcelHeaderDialog.Required.Column"),
				ColumnInfo.COLUMN_TYPE_CCOMBO, ReadExcelHeaderMeta.RequiredFilesDesc);
		colinfo[4] = new ColumnInfo(Messages.getString("ReadExcelHeaderDialog.IncludeSubDirs.Column"),
				ColumnInfo.COLUMN_TYPE_CCOMBO, ReadExcelHeaderMeta.RequiredFilesDesc);

		colinfo[0].setUsingVariables(true);
		colinfo[1].setUsingVariables(true);
		colinfo[1].setToolTip(Messages.getString("ReadExcelHeaderDialog.Files.Wildcard.Tooltip"));
		colinfo[2].setUsingVariables(true);
		colinfo[2].setToolTip(Messages.getString("ReadExcelHeaderDialog.Files.ExcludeWildcard.Tooltip"));

		wFilenameList = new TableView(transMeta, wFileComp, SWT.FULL_SELECTION | SWT.SINGLE | SWT.BORDER, colinfo, 2,
				lsMod, props);
		props.setLook(wFilenameList);

		fdFilenameList = new FormData();
		fdFilenameList.left = new FormAttachment(middle, 0);
		fdFilenameList.right = new FormAttachment(wbdFilename, -margin);
		fdFilenameList.top = new FormAttachment(wExcludeFilemask, margin);
		fdFilenameList.bottom = new FormAttachment(wbShowFiles, -margin);
		wFilenameList.setLayoutData(fdFilenameList);

		fdFileComp = new FormData();
		fdFileComp.left = new FormAttachment(0, 0);
		fdFileComp.top = new FormAttachment(0, 0);
		fdFileComp.right = new FormAttachment(100, 0);
		fdFileComp.bottom = new FormAttachment(100, 0);
		wFileComp.setLayoutData(fdFileComp);

		wFileComp.layout();
		wFileTab.setControl(wFileComp);

		// ///////////////////////////////////////////////////////////
		// / END OF FILE TAB
		// ///////////////////////////////////////////////////////////

		// ////////////////////////
		// START OF CONTENT TAB///
		// ////////////////////////
		wContentTab = new CTabItem(wTabFolder, SWT.NONE);
		wContentTab.setText(Messages.getString("ReadExcelHeaderDialog.Content.Tab"));

		FormLayout contentLayout = new FormLayout();
		contentLayout.marginWidth = 3;
		contentLayout.marginHeight = 3;

		wContentComp = new Composite(wTabFolder, SWT.NONE);
		props.setLook(wContentComp);
		wContentComp.setLayout(contentLayout);

		// /////////////////////////////////
		// START OF START ROW GROUP
		// /////////////////////////////////

		wStartRowGroup = new Group(wContentComp, SWT.SHADOW_NONE);
		props.setLook(wStartRowGroup);
		wStartRowGroup.setText(Messages.getString("ReadExcelHeaderDialog.Group.StartRowGroup.Label"));

		FormLayout startrowgroupLayout = new FormLayout();
		startrowgroupLayout.marginWidth = 10;
		startrowgroupLayout.marginHeight = 10;
		wStartRowGroup.setLayout(startrowgroupLayout);

		///////

		wlStartRowField = new Label(wStartRowGroup, SWT.RIGHT);
		wlStartRowField.setText(Messages.getString("ReadExcelHeaderDialog.StartRowField.Label"));
		props.setLook(wlStartRowField);
		fdlStartRowField = new FormData();
		fdlStartRowField.left = new FormAttachment(0, -margin);
		fdlStartRowField.top = new FormAttachment(0, margin);
		fdlStartRowField.right = new FormAttachment(middle, -2 * margin);
		wlStartRowField.setLayoutData(fdlStartRowField);

		wStartRowField = new Button(wStartRowGroup, SWT.CHECK);
		props.setLook(wStartRowField);
		wStartRowField.setToolTipText(Messages.getString("ReadExcelHeaderDialog.StartRowField.Tooltip"));
		fdStartRowField = new FormData();
		fdStartRowField.left = new FormAttachment(middle, -margin);
		fdStartRowField.top = new FormAttachment(0, margin);
		wStartRowField.setLayoutData(fdStartRowField);
		SelectionAdapter lStartRowfield = new SelectionAdapter() {
			public void widgetSelected(SelectionEvent arg0) {
				ActiveStartRowField();
				meta.setChanged(true);
			}
		};
		wStartRowField.addSelectionListener(lStartRowfield);

		// StartRow field
		wlStartRowSelField = new Label(wStartRowGroup, SWT.RIGHT);
		wlStartRowSelField.setText(Messages.getString("ReadExcelHeaderDialog.StartRowSelField.Label"));
		props.setLook(wlStartRowSelField);
		fdlStartRowSelField = new FormData();
		fdlStartRowSelField.left = new FormAttachment(0, -margin);
		fdlStartRowSelField.top = new FormAttachment(wStartRowField, margin);
		fdlStartRowSelField.right = new FormAttachment(middle, -2 * margin);
		wlStartRowSelField.setLayoutData(fdlStartRowSelField);

		wStartRowSelField = new CCombo(wStartRowGroup, SWT.BORDER | SWT.READ_ONLY);
		wStartRowSelField.setEditable(true);
		props.setLook(wStartRowSelField);
		wStartRowSelField.addModifyListener(lsMod);
		fdStartRowSelField = new FormData();
		fdStartRowSelField.left = new FormAttachment(middle, -margin);
		fdStartRowSelField.top = new FormAttachment(wStartRowField, margin);
		fdStartRowSelField.right = new FormAttachment(100, -margin);
		wStartRowSelField.setLayoutData(fdStartRowSelField);
		wStartRowSelField.addFocusListener(new FocusListener() {
			public void focusLost(org.eclipse.swt.events.FocusEvent e) {
			}

			public void focusGained(org.eclipse.swt.events.FocusEvent e) {
				Cursor busy = new Cursor(shell.getDisplay(), SWT.CURSOR_WAIT);
				shell.setCursor(busy);
				setStartRowField();
				shell.setCursor(null);
				busy.dispose();
			}
		});

		//////////

		wSeparator = new Label(wStartRowGroup, SWT.SEPARATOR | SWT.HORIZONTAL);
		fdSeparator = new FormData();
		fdSeparator.left = new FormAttachment(0, 0);
		fdSeparator.right = new FormAttachment(100, 0);
		fdSeparator.top = new FormAttachment(wStartRowSelField, margin);
		wSeparator.setLayoutData(fdSeparator);

		// start row line
		wLabelStepStartRow = new Label(wStartRowGroup, SWT.RIGHT);
		wLabelStepStartRow.setText(Messages.getString("ReadExcelHeaderDialog.StartRow.Label"));
		props.setLook(wLabelStepStartRow);
		wFormLabelStepStartRow = new FormData();
		wFormLabelStepStartRow.left = new FormAttachment(0, 0);
		wFormLabelStepStartRow.right = new FormAttachment(middle, -margin);
		wFormLabelStepStartRow.top = new FormAttachment(wSeparator, margin);
		wLabelStepStartRow.setLayoutData(wFormLabelStepStartRow);

		wTextStartRow = new TextVar(transMeta, wStartRowGroup, SWT.SINGLE | SWT.LEFT | SWT.BORDER);

		props.setLook(wTextStartRow);
		wTextStartRow.addModifyListener(lsMod);
		wFormStepStartRow = new FormData();
		wFormStepStartRow.left = new FormAttachment(wLabelStepStartRow, margin);
		wFormStepStartRow.top = new FormAttachment(wSeparator, margin);
		wFormStepStartRow.right = new FormAttachment(100, -margin);
		wTextStartRow.setLayoutData(wFormStepStartRow);

		fdStartRow = new FormData();
		fdStartRow.left = new FormAttachment(0, margin);
		fdStartRow.top = new FormAttachment(wSeparator, margin);
		fdStartRow.right = new FormAttachment(100, -margin);
		wStartRowGroup.setLayoutData(fdStartRow);

		// ///////////////////////////////////////////////////////////
		// / END OF START ROW GROUP
		// ///////////////////////////////////////////////////////////

		// /////////////////////////////////
		// START OF SAMPLE ROWS GROUP
		// /////////////////////////////////

		wSampleRowsGroup = new Group(wContentComp, SWT.SHADOW_NONE);
		props.setLook(wSampleRowsGroup);
		wSampleRowsGroup.setText(Messages.getString("ReadExcelHeaderDialog.Group.SampleRows.Label"));

		FormLayout samplerowsgroupLayout = new FormLayout();
		samplerowsgroupLayout.marginWidth = 10;
		samplerowsgroupLayout.marginHeight = 10;
		wSampleRowsGroup.setLayout(samplerowsgroupLayout);
		// sample rows line
		wLabelStepSampleRows = new Label(wSampleRowsGroup, SWT.RIGHT);
		wLabelStepSampleRows.setText(Messages.getString("ReadExcelHeaderDialog.SampleRows.Label"));
		props.setLook(wLabelStepSampleRows);
		wFormLabelStepSampleRows = new FormData();
		wFormLabelStepSampleRows.left = new FormAttachment(0, 0);
		wFormLabelStepSampleRows.right = new FormAttachment(middle, -margin);
		wFormLabelStepSampleRows.top = new FormAttachment(wStartRowGroup, margin);
		wLabelStepSampleRows.setLayoutData(wFormLabelStepSampleRows);

		wTextSampleRows = new TextVar(transMeta, wSampleRowsGroup, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wTextSampleRows);
		wTextSampleRows.addModifyListener(lsMod);
		wFormStepSampleRows = new FormData();
		wFormStepSampleRows.left = new FormAttachment(wLabelStepSampleRows, margin);
		wFormStepSampleRows.top = new FormAttachment(wStartRowGroup, margin);
		wFormStepSampleRows.right = new FormAttachment(100, -margin);
		wTextSampleRows.setLayoutData(wFormStepSampleRows);
		fdSampleRows = new FormData();
		fdSampleRows.left = new FormAttachment(0, margin);
		fdSampleRows.top = new FormAttachment(wStartRowGroup, margin);
		fdSampleRows.right = new FormAttachment(100, -margin);
		wSampleRowsGroup.setLayoutData(fdSampleRows);

		// /////////////////////////////////
		// END OF SAMPLE ROWS GROUP
		// /////////////////////////////////

		// do not fail if no files?
		wldoNotFailIfNoFile = new Label(wContentComp, SWT.RIGHT);
		wldoNotFailIfNoFile.setText(Messages.getString("ReadExcelHeaderDialog.doNotFailIfNoFile.Label"));
		props.setLook(wldoNotFailIfNoFile);
		fdldoNotFailIfNoFile = new FormData();
		fdldoNotFailIfNoFile.left = new FormAttachment(0, 0);
		fdldoNotFailIfNoFile.top = new FormAttachment(wSampleRowsGroup, 2 * margin);
		fdldoNotFailIfNoFile.right = new FormAttachment(middle, -margin);
		wldoNotFailIfNoFile.setLayoutData(fdldoNotFailIfNoFile);
		wdoNotFailIfNoFile = new Button(wContentComp, SWT.CHECK);
		props.setLook(wdoNotFailIfNoFile);
		wdoNotFailIfNoFile.setToolTipText(Messages.getString("ReadExcelHeaderDialog.doNotFailIfNoFile.Tooltip"));
		fddoNotFailIfNoFile = new FormData();
		fddoNotFailIfNoFile.left = new FormAttachment(middle, 0);
		fddoNotFailIfNoFile.top = new FormAttachment(wSampleRowsGroup, 2 * margin);
		wdoNotFailIfNoFile.setLayoutData(fddoNotFailIfNoFile);
		wdoNotFailIfNoFile.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent selectionEvent) {
				meta.setChanged();
			}
		});

		fdContentComp = new FormData();
		fdContentComp.left = new FormAttachment(0, 0);
		fdContentComp.top = new FormAttachment(0, 0);
		fdContentComp.right = new FormAttachment(100, 0);
		fdContentComp.bottom = new FormAttachment(100, 0);
		wContentComp.setLayoutData(fdContentComp);

		wContentComp.layout();
		wContentTab.setControl(wContentComp);

		// ///////////////////////////////////////////////////////////
		// / END OF CONTENT TAB
		// ///////////////////////////////////////////////////////////

		fdTabFolder = new FormData();
		fdTabFolder.left = new FormAttachment(0, 0);
		fdTabFolder.top = new FormAttachment(wStepname, margin);
		fdTabFolder.right = new FormAttachment(100, 0);
		fdTabFolder.bottom = new FormAttachment(100, -50);
		wTabFolder.setLayoutData(fdTabFolder);

		BaseStepDialog.positionBottomButtons(shell, new Button[] { wOK, wCancel }, margin, wTabFolder);

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
		wTextStartRow.addSelectionListener(lsDef);
		wTextSampleRows.addSelectionListener(lsDef);

		// Add the file to the list of files...
		SelectionAdapter selA = new SelectionAdapter() {
			public void widgetSelected(SelectionEvent arg0) {
				wFilenameList.add(new String[] { wFilename.getText(), wFilemask.getText(), wExcludeFilemask.getText(),
						ReadExcelHeaderMeta.RequiredFilesCode[0], ReadExcelHeaderMeta.RequiredFilesCode[0] });
				wFilename.setText("");
				wFilemask.setText("");
				wExcludeFilemask.setText("");
				wFilenameList.removeEmptyRows();
				wFilenameList.setRowNums();
				wFilenameList.optWidth(true);
			}
		};
		wbaFilename.addSelectionListener(selA);
		wFilename.addSelectionListener(selA);

		// Delete files from the list of files...
		wbdFilename.addSelectionListener(new SelectionAdapter() {
			public void widgetSelected(SelectionEvent arg0) {
				int[] idx = wFilenameList.getSelectionIndices();
				wFilenameList.remove(idx);
				wFilenameList.removeEmptyRows();
				wFilenameList.setRowNums();
				meta.setChanged();
			}
		});

		// Edit the selected file & remove from the list...
		wbeFilename.addSelectionListener(new SelectionAdapter() {
			public void widgetSelected(SelectionEvent arg0) {
				int idx = wFilenameList.getSelectionIndex();
				if (idx >= 0) {
					String[] string = wFilenameList.getItem(idx);
					wFilename.setText(string[0]);
					wFilemask.setText(string[1]);
					wExcludeFilemask.setText(string[2]);
					wFilenameList.remove(idx);
				}
				wFilenameList.removeEmptyRows();
				wFilenameList.setRowNums();
				meta.setChanged();
			}
		});

		// Show the files that are selected at this time...
		wbShowFiles.addSelectionListener(new SelectionAdapter() {
			public void widgetSelected(SelectionEvent e) {
				try {
					getInfo();
					FileInputList fileInputList = meta.getFiles(transMeta);
					String[] files = fileInputList.getFileStrings();

					if (files.length > 0) {
						EnterSelectionDialog esd = new EnterSelectionDialog(shell, files,
								Messages.getString("ReadExcelHeaderDialog.FilesReadSelection.DialogTitle"),
								Messages.getString("ReadExcelHeaderDialog.FilesReadSelection.DialogMessage"));
						esd.setViewOnly();
						esd.open();
					} else {
						MessageBox mb = new MessageBox(shell, SWT.OK | SWT.ICON_ERROR);
						mb.setMessage(Messages.getString("ReadExcelHeaderDialog.NoFileFound.DialogMessage"));
						mb.setText(Messages.getString("System.Dialog.Error.Title"));
						mb.open();
					}
				} catch (KettleException ex) {
					new ErrorDialog(shell, Messages.getString("ReadExcelHeaderDialog.ErrorParsingData.DialogTitle"),
							Messages.getString("ReadExcelHeaderDialog.ErrorParsingData.DialogMessage"), ex);
				}
			}
		});

		// Whenever something changes, set the tooltip to the expanded version of the
		// filename:
		wFilename.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				wFilename.setToolTipText(""); // StringUtil.environmentSubstitute( wFilename.getText() ) );
			}
		});

		// Listen to the Browse... button
		wbbFilename.addSelectionListener(
				new SelectionAdapterFileDialogTextVar(log, wFilename, transMeta, new SelectionAdapterOptions(
						SelectionOperation.FILE_OR_FOLDER, new FilterType[] { FilterType.ALL }, FilterType.ALL)));

		// Detect X or ALT-F4 or something that kills this window and cancel the dialog
		// properly
		shell.addShellListener(new ShellAdapter() {
			public void shellClosed(ShellEvent e) {
				cancel();
			}
		});

		wTabFolder.setSelection(0);

		// Set/Restore the dialog size based on last position on screen
		// The setSize() method is inherited from BaseStepDialog
		setSize();

		// restore the changed flag to original value, as the modify listeners fire
		// during dialog population
		getData();

		ActiveFileField();
		ActiveStartRowField();

		meta.setChanged(changed);

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

	private void ActiveFileField() {
		wlFilenameField.setEnabled(wFileField.getSelection());
		wFilenameField.setEnabled(wFileField.getSelection());
		wlWildcardField.setEnabled(wFileField.getSelection());
		wWildcardField.setEnabled(wFileField.getSelection());
		wlExcludeWildcardField.setEnabled(wFileField.getSelection());
		wExcludeWildcardField.setEnabled(wFileField.getSelection());
		wlIncludeSubFolder.setEnabled(wFileField.getSelection());
		wIncludeSubFolder.setEnabled(wFileField.getSelection());

		wlFilename.setEnabled(!wFileField.getSelection());
		wbbFilename.setEnabled(!wFileField.getSelection());
		wbaFilename.setEnabled(!wFileField.getSelection());
		wFilename.setEnabled(!wFileField.getSelection());
		wlFilemask.setEnabled(!wFileField.getSelection());
		wFilemask.setEnabled(!wFileField.getSelection());
		wlExcludeFilemask.setEnabled(!wFileField.getSelection());
		wExcludeFilemask.setEnabled(!wFileField.getSelection());
		wlFilenameList.setEnabled(!wFileField.getSelection());
		wbdFilename.setEnabled(!wFileField.getSelection());
		wbeFilename.setEnabled(!wFileField.getSelection());
		wbShowFiles.setEnabled(!wFileField.getSelection());
		wlFilenameList.setEnabled(!wFileField.getSelection());
		wFilenameList.setEnabled(!wFileField.getSelection());
	}

	private void ActiveStartRowField() {
		wlStartRowField.setEnabled(wStartRowField.getSelection());
		wStartRowSelField.setEnabled(wStartRowField.getSelection());

		wLabelStepStartRow.setEnabled(!wStartRowField.getSelection());
		wTextStartRow.setEnabled(!wStartRowField.getSelection());
	}

	private void setFileField() {
		try {
			if (!getpreviousFields) {
				getpreviousFields = true;
				String filename = wFilenameField.getText();
				String wildcard = wWildcardField.getText();
				String excludewildcard = wExcludeWildcardField.getText();

				wFilenameField.removeAll();
				wWildcardField.removeAll();
				wExcludeWildcardField.removeAll();

				RowMetaInterface r = transMeta.getPrevStepFields(stepname);
				if (r != null) {
					wFilenameField.setItems(r.getFieldNames());
					wWildcardField.setItems(r.getFieldNames());
					wExcludeWildcardField.setItems(r.getFieldNames());
				}
				if (filename != null) {
					wFilenameField.setText(filename);
				}
				if (wildcard != null) {
					wWildcardField.setText(wildcard);
				}
				if (excludewildcard != null) {
					wExcludeWildcardField.setText(excludewildcard);
				}
			}
		} catch (KettleException ke) {
			new ErrorDialog(shell, Messages.getString("GetFilesRowsCountDialog.FailedToGetFields.DialogTitle"),
					Messages.getString("GetFilesRowsCountDialog.FailedToGetFields.DialogMessage"), ke);
		}
	}

	private void setStartRowField() {
		try {

			wStartRowSelField.removeAll();

			RowMetaInterface r = transMeta.getPrevStepFields(stepname);
			if (r != null) {
				r.getFieldNames();

				for (int i = 0; i < r.getFieldNames().length; i++) {
					wStartRowSelField.add(r.getFieldNames()[i]);

				}
			}

		} catch (KettleException ke) {
			new ErrorDialog(shell, Messages.getString("GetFilesRowsCountDialog.FailedToGetFields.DialogTitle"),
					Messages.getString("GetFilesRowsCountDialog.FailedToGetFields.DialogMessage"), ke);
		}
	}

	/**
	 * Read the data from the ReadExcelHeaderMeta object and show it in this dialog.
	 */
	public void getData() {

		wTextSampleRows.setText(meta.getSampleRows());

		if (meta.getFileName() != null) {
			wFilenameList.removeAll();
			for (int i = 0; i < meta.getFileName().length; i++) {
				wFilenameList.add(new String[] { meta.getFileName()[i], meta.getFileMask()[i],
						meta.getExcludeFileMask()[i], meta.getRequiredFilesDesc(meta.getFileRequired()[i]),
						meta.getRequiredFilesDesc(meta.getIncludeSubFolders()[i]) });
			}
			wFilenameList.removeEmptyRows();
			wFilenameList.setRowNums();
			wFilenameList.optWidth(true);
		}

		wdoNotFailIfNoFile.setSelection(meta.isdoNotFailIfNoFile());

		wFileField.setSelection(meta.isFileField());
		wStartRowField.setSelection(meta.isStartRowField());
		if (meta.getFilenameField() != null) {
			wFilenameField.setText(meta.getFilenameField());
		}
		if (meta.getDynamicWildcardField() != null) {
			wWildcardField.setText(meta.getDynamicWildcardField());
		}
		if (meta.getDynamicExcludeWildcardField() != null) {
			wExcludeWildcardField.setText(meta.getDynamicExcludeWildcardField());
		}
		wIncludeSubFolder.setSelection(meta.isDynamicIncludeSubFolders());
		if (meta.isStartRowField()) {
			wStartRowSelField.setText(meta.getStartRowFieldName());
		} else {
			wTextStartRow.setText(meta.getStartRow());
		}

		logDebug("Finished getting step data");

		wStepname.selectAll();
		wStepname.setFocus();
	}

	private void getInfo() throws KettleException {
		stepname = wStepname.getText(); // return value
		int nrFiles = wFilenameList.getItemCount();
		meta.allocate(nrFiles);
		meta.setFileName(wFilenameList.getItems(0));
		meta.setFileMask(wFilenameList.getItems(1));
		meta.setExcludeFileMask(wFilenameList.getItems(2));
		meta.setFileRequired(wFilenameList.getItems(3));
		meta.setIncludeSubFolders(wFilenameList.getItems(4));

		meta.setSampleRows(wTextSampleRows.getText());

		meta.setStartRowField(wStartRowField.getSelection());
		meta.setStartRowFieldName(wStartRowSelField.getText());
		meta.setStartRow(wTextStartRow.getText());

		meta.setFileField(wFileField.getSelection());
		meta.setFileNameField(wFilenameField.getText());
		meta.setDynamicWildcardField( wWildcardField.getText() );
    	meta.setDynamicExcludeWildcardField( wExcludeWildcardField.getText() );
		meta.setDynamicIncludeSubFolders( wIncludeSubFolder.getSelection() );
    	meta.setdoNotFailIfNoFile( wdoNotFailIfNoFile.getSelection() );
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

		if (Utils.isEmpty(wStepname.getText())) {
			return;
		}

		try {
			getInfo();
		} catch (KettleException e) {
			new ErrorDialog(shell, Messages.getString("ReadExcelHeaderDialog.ErrorParsingData.DialogTitle"),
					Messages.getString("ReadExcelHeaderDialog.ErrorParsingData.DialogMessage"), e);
		}

		meta.setChanged();
		// close the SWT dialog window
		dispose();
	}
}
