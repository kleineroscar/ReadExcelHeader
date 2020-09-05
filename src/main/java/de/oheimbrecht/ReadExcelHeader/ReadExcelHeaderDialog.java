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

import org.apache.hop.core.Const;
import org.apache.hop.core.Props;
import org.apache.hop.core.exception.HopException;
import org.apache.hop.core.row.IRowMeta;
import org.apache.hop.core.util.Utils;
import org.apache.hop.i18n.BaseMessages;
import org.apache.hop.pipeline.Pipeline;
import org.apache.hop.pipeline.PipelineMeta;
import org.apache.hop.pipeline.PipelinePreviewFactory;
import org.apache.hop.pipeline.transform.BaseTransformMeta;
import org.apache.hop.pipeline.transform.ITransformDialog;
import org.apache.hop.ui.core.dialog.*;
import org.apache.hop.ui.core.widget.ColumnInfo;
import org.apache.hop.ui.core.widget.TableView;
import org.apache.hop.ui.core.widget.TextVar;
import org.apache.hop.ui.pipeline.dialog.PipelinePreviewProgressDialog;
import org.apache.hop.ui.pipeline.transform.BaseTransformDialog;
import org.eclipse.swt.SWT;
import org.eclipse.swt.custom.CCombo;
import org.eclipse.swt.custom.CTabFolder;
import org.eclipse.swt.custom.CTabItem;
import org.eclipse.swt.events.*;
import org.eclipse.swt.graphics.Cursor;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.widgets.*;

public class ReadExcelHeaderDialog extends BaseTransformDialog implements ITransformDialog {
	private static Class<?> PKG = ReadExcelHeader.class; // for i18n purposes, needed by Translator!!
private static final String[] YES_NO_COMBO = new String[] {
    BaseMessages.getString( PKG, "System.Combo.No" ), BaseMessages.getString( PKG, "System.Combo.Yes" ) };

  // do not fail if no files?
  private Label wldoNotFailIfNoFile;
  private Button wdoNotFailIfNoFile;
  private FormData fdldoNotFailIfNoFile, fddoNotFailIfNoFile;

  private CTabFolder wContentFolder;

  private FormData fdTabFolder;

  private CTabItem wFileTab, wContentTab;

  private Composite wFileComp, wContentComp;

  private FormData fdFileComp, fdContentComp;

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

  private Label wlExcludeFilemask;

  private TextVar wExcludeFilemask;

  private FormData fdlExcludeFilemask, fdExcludeFilemask;

  private Label wlFilemask;

  private TextVar wFilemask;

  private FormData fdlFilemask, fdFilemask;

  private Button wbShowFiles;

  private FormData fdbShowFiles;

  private ReadExcelHeaderMeta input;

  private int middle, margin;

  private ModifyListener lsMod;

  private Group wOriginFiles;

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

  private Group wStartRowGroup;
  private Label wLabelStepStartRow;
  private FormData wFormLabelStepStartRow, wFormStepStartRow, fdStartRow;
  private TextVar wTextStartRow;
  private Label wlStartRowField;
	private FormData fdlStartRowField;
	private Button wStartRowField;
  private FormData fdStartRowField;
  private Label wlStartRowSelField;
	private FormData fdlStartRowSelField;
	private CCombo wStartRowSelField;
	private FormData fdStartRowSelField;
  
  private Group wSampleRowsGroup;
  private Label wLabelStepSampleRows;
  private FormData wFormLabelStepSampleRows, wFormStepSampleRows, fdSampleRows;
  private TextVar wTextSampleRows;
  
	private Label wSeparator;
	private FormData fdSeparator;

  private boolean getpreviousFields = false;

  public ReadExcelHeaderDialog( Shell parent, Object in, PipelineMeta pipelineMeta, String sname ) {
    super( parent, (BaseTransformMeta) in, pipelineMeta, sname );
    input = (ReadExcelHeaderMeta) in;
  }

  public String open() {
    Shell parent = getParent();
    Display display = parent.getDisplay();

    shell = new Shell( parent, SWT.DIALOG_TRIM | SWT.RESIZE | SWT.MAX | SWT.MIN );
    props.setLook( shell );
    setShellImage( shell, input );

    lsMod = new ModifyListener() {
      public void modifyText( ModifyEvent e ) {
        input.setChanged();
      }
    };
    changed = input.hasChanged();

    FormLayout formLayout = new FormLayout();
    formLayout.marginWidth = Const.FORM_MARGIN;
    formLayout.marginHeight = Const.FORM_MARGIN;

    shell.setLayout( formLayout );
    shell.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.DialogTitle" ) );

    middle = props.getMiddlePct();
    margin = props.getMargin();

    // TransformName line
    wlTransformName = new Label( shell, SWT.RIGHT );
    wlTransformName.setText( BaseMessages.getString( PKG, "System.Label.TransformName" ) );
    props.setLook( wlTransformName );
    fdlTransformName = new FormData();
    fdlTransformName.left = new FormAttachment( 0, 0 );
    fdlTransformName.top = new FormAttachment( 0, margin );
    fdlTransformName.right = new FormAttachment( middle, -margin );
    wlTransformName.setLayoutData( fdlTransformName );
    wTransformName = new Text( shell, SWT.SINGLE | SWT.LEFT | SWT.BORDER );
    wTransformName.setText( transformName );
    props.setLook( wTransformName );
    wTransformName.addModifyListener( lsMod );
    fdTransformName = new FormData();
    fdTransformName.left = new FormAttachment( middle, 0 );
    fdTransformName.top = new FormAttachment( 0, margin );
    fdTransformName.right = new FormAttachment( 100, 0 );
    wTransformName.setLayoutData( fdTransformName );

    wContentFolder = new CTabFolder( shell, SWT.BORDER );
    props.setLook( wContentFolder, Props.WIDGET_STYLE_TAB );

    // ////////////////////////
    // START OF FILE TAB ///
    // ////////////////////////
    wFileTab = new CTabItem( wContentFolder, SWT.NONE );
    wFileTab.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FileTab.TabTitle" ) );

    wFileComp = new Composite( wContentFolder, SWT.NONE );
    props.setLook( wFileComp );

    FormLayout fileLayout = new FormLayout();
    fileLayout.marginWidth = 3;
    fileLayout.marginHeight = 3;
    wFileComp.setLayout( fileLayout );

    // ///////////////////////////////
    // START OF Origin files GROUP //
    // ///////////////////////////////

    wOriginFiles = new Group( wFileComp, SWT.SHADOW_NONE );
    props.setLook( wOriginFiles );
    wOriginFiles.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.wOriginFiles.Label" ) );

    FormLayout OriginFilesgroupLayout = new FormLayout();
    OriginFilesgroupLayout.marginWidth = 10;
    OriginFilesgroupLayout.marginHeight = 10;
    wOriginFiles.setLayout( OriginFilesgroupLayout );

    // Is Filename defined in a Field
    wlFileField = new Label( wOriginFiles, SWT.RIGHT );
    wlFileField.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FileField.Label" ) );
    props.setLook( wlFileField );
    fdlFileField = new FormData();
    fdlFileField.left = new FormAttachment( 0, -margin );
    fdlFileField.top = new FormAttachment( 0, margin );
    fdlFileField.right = new FormAttachment( middle, -2 * margin );
    wlFileField.setLayoutData( fdlFileField );

    wFileField = new Button( wOriginFiles, SWT.CHECK );
    props.setLook( wFileField );
    wFileField.setToolTipText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FileField.Tooltip" ) );
    fdFileField = new FormData();
    fdFileField.left = new FormAttachment( middle, -margin );
    fdFileField.top = new FormAttachment( 0, margin );
    wFileField.setLayoutData( fdFileField );
    SelectionAdapter lfilefield = new SelectionAdapter() {
      public void widgetSelected( SelectionEvent arg0 ) {
        ActiveFileField();
        setFileField();
        input.setChanged();
      }
    };
    wFileField.addSelectionListener( lfilefield );

    // Filename field
    wlFilenameField = new Label( wOriginFiles, SWT.RIGHT );
    wlFilenameField.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.wlFilenameField.Label" ) );
    props.setLook( wlFilenameField );
    fdlFilenameField = new FormData();
    fdlFilenameField.left = new FormAttachment( 0, -margin );
    fdlFilenameField.top = new FormAttachment( wFileField, margin );
    fdlFilenameField.right = new FormAttachment( middle, -2 * margin );
    wlFilenameField.setLayoutData( fdlFilenameField );

    wFilenameField = new CCombo( wOriginFiles, SWT.BORDER | SWT.READ_ONLY );
    wFilenameField.setEditable( true );
    props.setLook( wFilenameField );
    wFilenameField.addModifyListener( lsMod );
    fdFilenameField = new FormData();
    fdFilenameField.left = new FormAttachment( middle, -margin );
    fdFilenameField.top = new FormAttachment( wFileField, margin );
    fdFilenameField.right = new FormAttachment( 100, -margin );
    wFilenameField.setLayoutData( fdFilenameField );

    // Wildcard field
    wlWildcardField = new Label( wOriginFiles, SWT.RIGHT );
    wlWildcardField.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.wlWildcardField.Label" ) );
    props.setLook( wlWildcardField );
    fdlWildcardField = new FormData();
    fdlWildcardField.left = new FormAttachment( 0, -margin );
    fdlWildcardField.top = new FormAttachment( wFilenameField, margin );
    fdlWildcardField.right = new FormAttachment( middle, -2 * margin );
    wlWildcardField.setLayoutData( fdlWildcardField );

    wWildcardField = new CCombo( wOriginFiles, SWT.BORDER | SWT.READ_ONLY );
    wWildcardField.setEditable( true );
    props.setLook( wWildcardField );
    wWildcardField.addModifyListener( lsMod );
    fdWildcardField = new FormData();
    fdWildcardField.left = new FormAttachment( middle, -margin );
    fdWildcardField.top = new FormAttachment( wFilenameField, margin );
    fdWildcardField.right = new FormAttachment( 100, -margin );
    wWildcardField.setLayoutData( fdWildcardField );

    // ExcludeWildcard field
    wlExcludeWildcardField = new Label( wOriginFiles, SWT.RIGHT );
    wlExcludeWildcardField.setText( BaseMessages
      .getString( PKG, "ReadExcelHeaderDialog.wlExcludeWildcardField.Label" ) );
    props.setLook( wlExcludeWildcardField );
    fdlExcludeWildcardField = new FormData();
    fdlExcludeWildcardField.left = new FormAttachment( 0, -margin );
    fdlExcludeWildcardField.top = new FormAttachment( wWildcardField, margin );
    fdlExcludeWildcardField.right = new FormAttachment( middle, -2 * margin );
    wlExcludeWildcardField.setLayoutData( fdlExcludeWildcardField );

    wExcludeWildcardField = new CCombo( wOriginFiles, SWT.BORDER | SWT.READ_ONLY );
    wExcludeWildcardField.setEditable( true );
    props.setLook( wExcludeWildcardField );
    wExcludeWildcardField.addModifyListener( lsMod );
    fdExcludeWildcardField = new FormData();
    fdExcludeWildcardField.left = new FormAttachment( middle, -margin );
    fdExcludeWildcardField.top = new FormAttachment( wWildcardField, margin );
    fdExcludeWildcardField.right = new FormAttachment( 100, -margin );
    wExcludeWildcardField.setLayoutData( fdExcludeWildcardField );

    // Is includeSubFoldername defined in a Field
    wlIncludeSubFolder = new Label( wOriginFiles, SWT.RIGHT );
    wlIncludeSubFolder.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.includeSubFolder.Label" ) );
    props.setLook( wlIncludeSubFolder );
    fdlIncludeSubFolder = new FormData();
    fdlIncludeSubFolder.left = new FormAttachment( 0, -margin );
    fdlIncludeSubFolder.top = new FormAttachment( wExcludeWildcardField, margin );
    fdlIncludeSubFolder.right = new FormAttachment( middle, -2 * margin );
    wlIncludeSubFolder.setLayoutData( fdlIncludeSubFolder );

    wIncludeSubFolder = new Button( wOriginFiles, SWT.CHECK );
    props.setLook( wIncludeSubFolder );
    wIncludeSubFolder
      .setToolTipText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.includeSubFolder.Tooltip" ) );
    fdIncludeSubFolder = new FormData();
    fdIncludeSubFolder.left = new FormAttachment( middle, -margin );
    fdIncludeSubFolder.top = new FormAttachment( wExcludeWildcardField, margin );
    wIncludeSubFolder.setLayoutData( fdIncludeSubFolder );
    wIncludeSubFolder.addSelectionListener( new SelectionAdapter() {
      @Override
      public void widgetSelected( SelectionEvent selectionEvent ) {
        input.setChanged();
      }
    } );

    fdOriginFiles = new FormData();
    fdOriginFiles.left = new FormAttachment( 0, margin );
    fdOriginFiles.top = new FormAttachment( wFilenameList, margin );
    fdOriginFiles.right = new FormAttachment( 100, -margin );
    wOriginFiles.setLayoutData( fdOriginFiles );

    // ///////////////////////////////////////////////////////////
    // / END OF Origin files GROUP
    // ///////////////////////////////////////////////////////////

    // Filename line
    wlFilename = new Label( wFileComp, SWT.RIGHT );
    wlFilename.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Filename.Label" ) );
    props.setLook( wlFilename );
    fdlFilename = new FormData();
    fdlFilename.left = new FormAttachment( 0, 0 );
    fdlFilename.top = new FormAttachment( wOriginFiles, margin );
    fdlFilename.right = new FormAttachment( middle, -margin );
    wlFilename.setLayoutData( fdlFilename );

    wbbFilename = new Button( wFileComp, SWT.PUSH | SWT.CENTER );
    props.setLook( wbbFilename );
    wbbFilename.setText( BaseMessages.getString( PKG, "System.Button.Browse" ) );
    wbbFilename.setToolTipText( BaseMessages.getString( PKG, "System.Tooltip.BrowseForFileOrDirAndAdd" ) );
    fdbFilename = new FormData();
    fdbFilename.right = new FormAttachment( 100, 0 );
    fdbFilename.top = new FormAttachment( wOriginFiles, margin );
    wbbFilename.setLayoutData( fdbFilename );

    wbaFilename = new Button( wFileComp, SWT.PUSH | SWT.CENTER );
    props.setLook( wbaFilename );
    wbaFilename.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameAdd.Button" ) );
    wbaFilename.setToolTipText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameAdd.Tooltip" ) );
    fdbaFilename = new FormData();
    fdbaFilename.right = new FormAttachment( wbbFilename, -margin );
    fdbaFilename.top = new FormAttachment( wOriginFiles, margin );
    wbaFilename.setLayoutData( fdbaFilename );

    wFilename = new TextVar( pipelineMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER );
    props.setLook( wFilename );
    wFilename.addModifyListener( lsMod );
    fdFilename = new FormData();
    fdFilename.left = new FormAttachment( middle, 0 );
    fdFilename.right = new FormAttachment( wbaFilename, -margin );
    fdFilename.top = new FormAttachment( wOriginFiles, margin );
    wFilename.setLayoutData( fdFilename );

    wlFilemask = new Label( wFileComp, SWT.RIGHT );
    wlFilemask.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Filemask.Label" ) );
    props.setLook( wlFilemask );
    fdlFilemask = new FormData();
    fdlFilemask.left = new FormAttachment( 0, 0 );
    fdlFilemask.top = new FormAttachment( wFilename, margin );
    fdlFilemask.right = new FormAttachment( middle, -margin );
    wlFilemask.setLayoutData( fdlFilemask );
    wFilemask = new TextVar( pipelineMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER );
    props.setLook( wFilemask );
    wFilemask.addModifyListener( lsMod );
    fdFilemask = new FormData();
    fdFilemask.left = new FormAttachment( middle, 0 );
    fdFilemask.top = new FormAttachment( wFilename, margin );
    fdFilemask.right = new FormAttachment( wFilename, 0, SWT.RIGHT );
    wFilemask.setLayoutData( fdFilemask );

    wlExcludeFilemask = new Label( wFileComp, SWT.RIGHT );
    wlExcludeFilemask.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.ExcludeFilemask.Label" ) );
    props.setLook( wlExcludeFilemask );
    fdlExcludeFilemask = new FormData();
    fdlExcludeFilemask.left = new FormAttachment( 0, 0 );
    fdlExcludeFilemask.top = new FormAttachment( wFilemask, margin );
    fdlExcludeFilemask.right = new FormAttachment( middle, -margin );
    wlExcludeFilemask.setLayoutData( fdlExcludeFilemask );
    wExcludeFilemask = new TextVar( pipelineMeta, wFileComp, SWT.SINGLE | SWT.LEFT | SWT.BORDER );
    props.setLook( wExcludeFilemask );
    wExcludeFilemask.addModifyListener( lsMod );
    fdExcludeFilemask = new FormData();
    fdExcludeFilemask.left = new FormAttachment( middle, 0 );
    fdExcludeFilemask.top = new FormAttachment( wFilemask, margin );
    fdExcludeFilemask.right = new FormAttachment( wFilename, 0, SWT.RIGHT );
    wExcludeFilemask.setLayoutData( fdExcludeFilemask );

    // Filename list line
    wlFilenameList = new Label( wFileComp, SWT.RIGHT );
    wlFilenameList.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameList.Label" ) );
    props.setLook( wlFilenameList );
    fdlFilenameList = new FormData();
    fdlFilenameList.left = new FormAttachment( 0, 0 );
    fdlFilenameList.top = new FormAttachment( wExcludeFilemask, margin );
    fdlFilenameList.right = new FormAttachment( middle, -margin );
    wlFilenameList.setLayoutData( fdlFilenameList );

    // Buttons to the right of the screen...
    wbdFilename = new Button( wFileComp, SWT.PUSH | SWT.CENTER );
    props.setLook( wbdFilename );
    wbdFilename.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameDelete.Button" ) );
    wbdFilename.setToolTipText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameDelete.Tooltip" ) );
    fdbdFilename = new FormData();
    fdbdFilename.right = new FormAttachment( 100, 0 );
    fdbdFilename.top = new FormAttachment( wExcludeFilemask, 40 );
    wbdFilename.setLayoutData( fdbdFilename );

    wbeFilename = new Button( wFileComp, SWT.PUSH | SWT.CENTER );
    props.setLook( wbeFilename );
    wbeFilename.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameEdit.Button" ) );
    wbeFilename.setToolTipText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FilenameEdit.Tooltip" ) );
    fdbeFilename = new FormData();
    fdbeFilename.right = new FormAttachment( 100, 0 );
    fdbeFilename.left = new FormAttachment( wbdFilename, 0, SWT.LEFT );
    fdbeFilename.top = new FormAttachment( wbdFilename, margin );
    wbeFilename.setLayoutData( fdbeFilename );

    wbShowFiles = new Button( wFileComp, SWT.PUSH | SWT.CENTER );
    props.setLook( wbShowFiles );
    wbShowFiles.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.ShowFiles.Button" ) );
    fdbShowFiles = new FormData();
    fdbShowFiles.left = new FormAttachment( middle, 0 );
    fdbShowFiles.bottom = new FormAttachment( 100, 0 );
    wbShowFiles.setLayoutData( fdbShowFiles );

    ColumnInfo[] colinfo =
      new ColumnInfo[] {
        new ColumnInfo(
          BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FileDirColumn.Column" ),
          ColumnInfo.COLUMN_TYPE_TEXT, false ),
        new ColumnInfo(
          BaseMessages.getString( PKG, "ReadExcelHeaderDialog.WildcardColumn.Column" ),
          ColumnInfo.COLUMN_TYPE_TEXT, false ),
        new ColumnInfo(
          BaseMessages.getString( PKG, "ReadExcelHeaderDialog.ExcludeWildcardColumn.Column" ),
          ColumnInfo.COLUMN_TYPE_TEXT, false ),
        new ColumnInfo(
          BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Required.Column" ),
          ColumnInfo.COLUMN_TYPE_CCOMBO, YES_NO_COMBO ),
        new ColumnInfo(
          BaseMessages.getString( PKG, "ReadExcelHeaderDialog.IncludeSubDirs.Column" ),
          ColumnInfo.COLUMN_TYPE_CCOMBO, YES_NO_COMBO ) };

    colinfo[ 0 ].setUsingVariables( true );
    colinfo[ 1 ].setUsingVariables( true );
    colinfo[ 1 ].setToolTip( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.WildcardColumn.Tooltip" ) );
    colinfo[ 2 ].setUsingVariables( true );
    colinfo[ 2 ].setToolTip( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.ExcludeWildcardColumn.Tooltip" ) );
    colinfo[ 3 ].setToolTip( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Required.Tooltip" ) );
    colinfo[ 4 ].setToolTip( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.IncludeSubDirs.ToolTip" ) );

    wFilenameList =
      new TableView(
        pipelineMeta, wFileComp, SWT.FULL_SELECTION | SWT.SINGLE | SWT.BORDER, colinfo, colinfo.length, lsMod,
        props );
    props.setLook( wFilenameList );
    fdFilenameList = new FormData();
    fdFilenameList.left = new FormAttachment( middle, 0 );
    fdFilenameList.right = new FormAttachment( wbdFilename, -margin );
    fdFilenameList.top = new FormAttachment( wExcludeFilemask, margin );
    fdFilenameList.bottom = new FormAttachment( wbShowFiles, -margin );
    wFilenameList.setLayoutData( fdFilenameList );

    fdFileComp = new FormData();
    fdFileComp.left = new FormAttachment( 0, 0 );
    fdFileComp.top = new FormAttachment( 0, 0 );
    fdFileComp.right = new FormAttachment( 100, 0 );
    fdFileComp.bottom = new FormAttachment( 100, 0 );
    wFileComp.setLayoutData( fdFileComp );

    wFileComp.layout();
    wFileTab.setControl( wFileComp );

    // ///////////////////////////////////////////////////////////
    // / END OF FILE TAB
    // ///////////////////////////////////////////////////////////

    fdTabFolder = new FormData();
    fdTabFolder.left = new FormAttachment( 0, 0 );
    fdTabFolder.top = new FormAttachment( wTransformName, margin );
    fdTabFolder.right = new FormAttachment( 100, 0 );
    fdTabFolder.bottom = new FormAttachment( 100, -50 );
    wContentFolder.setLayoutData( fdTabFolder );

    // ////////////////////////
    // START OF Content TAB ///
    // ////////////////////////
    wContentTab = new CTabItem( wContentFolder, SWT.NONE );
    wContentTab.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.ContentTab.TabTitle" ) );

    wContentComp = new Composite( wContentFolder, SWT.NONE );
    props.setLook( wContentComp );

    FormLayout filesettingLayout = new FormLayout();
    filesettingLayout.marginWidth = 3;
    filesettingLayout.marginHeight = 3;
    wContentComp.setLayout( fileLayout );

    // /////////////////////////////////
		// START OF START ROW GROUP
		// /////////////////////////////////

		wStartRowGroup = new Group(wContentComp, SWT.SHADOW_NONE );
		props.setLook(wStartRowGroup);
		wStartRowGroup.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Group.StartRowGroup.Label"));

		FormLayout startrowgroupLayout = new FormLayout();
		startrowgroupLayout.marginWidth = 10;
		startrowgroupLayout.marginHeight = 10;
		wStartRowGroup.setLayout( startrowgroupLayout );

		///////

		wlStartRowField = new Label(wStartRowGroup, SWT.RIGHT);
		wlStartRowField.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.StartRowField.Label"));
		props.setLook(wlStartRowField);
		fdlStartRowField = new FormData();
		fdlStartRowField.left = new FormAttachment(0, 0);
		fdlStartRowField.top = new FormAttachment(0, 3 * margin);
    fdlStartRowField.right = new FormAttachment(middle, -margin);

		wlStartRowField.setLayoutData(fdlStartRowField);

		wStartRowField = new Button(wStartRowGroup, SWT.CHECK);
		props.setLook(wStartRowField);
		wStartRowField.setToolTipText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.StartRowField.Tooltip"));
		fdStartRowField = new FormData();
		fdStartRowField.left = new FormAttachment( middle, 0 );
		fdStartRowField.top = new FormAttachment(0, 3 * margin );
		wStartRowField.setLayoutData(fdStartRowField);
		wStartRowField.addSelectionListener( new SelectionAdapter() {
      @Override
      public void widgetSelected( SelectionEvent selectionEvent ) {
        ActiveStartRowField();
        input.setChanged();
      }
    } );

		// StartRow field
		wlStartRowSelField = new Label(wStartRowGroup, SWT.RIGHT);
		wlStartRowSelField.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.StartRowSelField.Label"));
		props.setLook(wlStartRowSelField);
		fdlStartRowSelField = new FormData();
		fdlStartRowSelField.left = new FormAttachment(0, 0);
		fdlStartRowSelField.top = new FormAttachment(wStartRowField, 2 * margin);
		fdlStartRowSelField.right = new FormAttachment( middle, -margin );
		wlStartRowSelField.setLayoutData(fdlStartRowSelField);

		wStartRowSelField = new CCombo(wStartRowGroup, SWT.BORDER | SWT.READ_ONLY);
		wStartRowSelField.setEditable(true);
		props.setLook(wStartRowSelField);
		wStartRowSelField.addModifyListener(lsMod);
		fdStartRowSelField = new FormData();
		fdStartRowSelField.left = new FormAttachment(middle, 0);
		fdStartRowSelField.top = new FormAttachment(wStartRowField, 2 * margin);
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
		wLabelStepStartRow.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.StartRow.Label"));
		props.setLook(wLabelStepStartRow);
		wFormLabelStepStartRow = new FormData();
    wFormLabelStepStartRow.left = new FormAttachment(0, 0);
    wFormLabelStepStartRow.top = new FormAttachment(wSeparator, 2 * margin);
		wFormLabelStepStartRow.right = new FormAttachment( middle, -margin );
		wLabelStepStartRow.setLayoutData(wFormLabelStepStartRow);

		wTextStartRow = new TextVar(pipelineMeta, wStartRowGroup, SWT.SINGLE | SWT.LEFT | SWT.BORDER);

		props.setLook(wTextStartRow);
		wTextStartRow.addModifyListener(lsMod);
		wFormStepStartRow = new FormData();
		wFormStepStartRow.left = new FormAttachment(wLabelStepStartRow, 0);
		wFormStepStartRow.top = new FormAttachment(wSeparator, 2 * margin);
		wFormStepStartRow.right = new FormAttachment(100, -margin);
		wTextStartRow.setLayoutData(wFormStepStartRow);

		fdStartRow = new FormData();
		fdStartRow.left = new FormAttachment(0, 0);
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
		wSampleRowsGroup.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Group.SampleRows.Label"));

		FormLayout samplerowsgroupLayout = new FormLayout();
		samplerowsgroupLayout.marginWidth = 10;
		samplerowsgroupLayout.marginHeight = 10;
		wSampleRowsGroup.setLayout(samplerowsgroupLayout);
		// sample rows line
		wLabelStepSampleRows = new Label(wSampleRowsGroup, SWT.RIGHT);
		wLabelStepSampleRows.setText(BaseMessages.getString( PKG, "ReadExcelHeaderDialog.SampleRows.Label"));
		props.setLook(wLabelStepSampleRows);
		wFormLabelStepSampleRows = new FormData();
		wFormLabelStepSampleRows.left = new FormAttachment(0, 0);
		wFormLabelStepSampleRows.right = new FormAttachment(middle, -margin);
		wFormLabelStepSampleRows.top = new FormAttachment(wStartRowGroup, 2 * margin);
		wLabelStepSampleRows.setLayoutData(wFormLabelStepSampleRows);

		wTextSampleRows = new TextVar(pipelineMeta, wSampleRowsGroup, SWT.SINGLE | SWT.LEFT | SWT.BORDER);
		props.setLook(wTextSampleRows);
		wTextSampleRows.addModifyListener(lsMod);
		wFormStepSampleRows = new FormData();
		wFormStepSampleRows.left = new FormAttachment(wLabelStepSampleRows, margin);
		wFormStepSampleRows.top = new FormAttachment(wStartRowGroup, 2 * margin);
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
    wldoNotFailIfNoFile = new Label( wContentComp, SWT.RIGHT );
    wldoNotFailIfNoFile.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.doNotFailIfNoFile.Label" ) );
    props.setLook( wldoNotFailIfNoFile );
    fdldoNotFailIfNoFile = new FormData();
    fdldoNotFailIfNoFile.left = new FormAttachment( 0, 0 );
    fdldoNotFailIfNoFile.top = new FormAttachment( wSampleRowsGroup, 2 * margin );
    fdldoNotFailIfNoFile.right = new FormAttachment( middle, -margin );
    wldoNotFailIfNoFile.setLayoutData( fdldoNotFailIfNoFile );
    wdoNotFailIfNoFile = new Button( wContentComp, SWT.CHECK );
    props.setLook( wdoNotFailIfNoFile );
    wdoNotFailIfNoFile.setToolTipText( BaseMessages
      .getString( PKG, "ReadExcelHeaderDialog.doNotFailIfNoFile.Tooltip" ) );
    fddoNotFailIfNoFile = new FormData();
    fddoNotFailIfNoFile.left = new FormAttachment( middle, 0 );
    fddoNotFailIfNoFile.top = new FormAttachment( wSampleRowsGroup, 2 * margin );
    wdoNotFailIfNoFile.setLayoutData( fddoNotFailIfNoFile );
    wdoNotFailIfNoFile.addSelectionListener( new SelectionAdapter() {
      @Override
      public void widgetSelected( SelectionEvent selectionEvent ) {
        input.setChanged();
      }
    } );


    fdContentComp = new FormData();
    fdContentComp.left = new FormAttachment( 0, 0 );
    fdContentComp.top = new FormAttachment( 0, 0 );
    fdContentComp.right = new FormAttachment( 100, 0 );
    fdContentComp.bottom = new FormAttachment( 100, 0 );
    wContentComp.setLayoutData( fdContentComp );

    wContentComp.layout();
    wContentTab.setControl( wContentComp );

    // ///////////////////////////////////////////////////////////
    // / END OF Content TAB
    // ///////////////////////////////////////////////////////////

    wOk = new Button( shell, SWT.PUSH );
    wOk.setText( BaseMessages.getString( PKG, "System.Button.OK" ) );

    wPreview = new Button( shell, SWT.PUSH );
    wPreview.setText( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.Preview.Button" ) );

    wCancel = new Button( shell, SWT.PUSH );
    wCancel.setText( BaseMessages.getString( PKG, "System.Button.Cancel" ) );

    setButtonPositions( new Button[] { wOk, wPreview, wCancel }, margin, wContentFolder );

    // Add listeners
    lsOk = new Listener() {
      public void handleEvent( Event e ) {
        ok();
      }
    };
    lsPreview = new Listener() {
      public void handleEvent( Event e ) {
        preview();
      }
    };
    lsCancel = new Listener() {
      public void handleEvent( Event e ) {
        cancel();
      }
    };

    wOk.addListener( SWT.Selection, lsOk );
    wPreview.addListener( SWT.Selection, lsPreview );
    wCancel.addListener( SWT.Selection, lsCancel );

    lsDef = new SelectionAdapter() {
      public void widgetDefaultSelected( SelectionEvent e ) {
        ok();
      }
    };

    wTransformName.addSelectionListener( lsDef );

    // Add the file to the list of files...
    SelectionAdapter selA = new SelectionAdapter() {
      public void widgetSelected( SelectionEvent arg0 ) {
        wFilenameList.add( new String[] {
          wFilename.getText(), wFilemask.getText(), wExcludeFilemask.getText(),
          ReadExcelHeaderMeta.RequiredFilesCode[ 0 ], ReadExcelHeaderMeta.RequiredFilesCode[ 0 ] } );
        wFilename.setText( "" );
        wFilemask.setText( "" );
        wFilenameList.removeEmptyRows();
        wFilenameList.setRowNums();
        wFilenameList.optWidth( true );
      }
    };
    wbaFilename.addSelectionListener( selA );
    wFilename.addSelectionListener( selA );

    // Delete files from the list of files...
    wbdFilename.addSelectionListener( new SelectionAdapter() {
      public void widgetSelected( SelectionEvent arg0 ) {
        int[] idx = wFilenameList.getSelectionIndices();
        wFilenameList.remove( idx );
        wFilenameList.removeEmptyRows();
        wFilenameList.setRowNums();
        input.setChanged();
      }
    } );

    // Edit the selected file & remove from the list...
    wbeFilename.addSelectionListener( new SelectionAdapter() {
      public void widgetSelected( SelectionEvent arg0 ) {
        int idx = wFilenameList.getSelectionIndex();
        if ( idx >= 0 ) {
          String[] string = wFilenameList.getItem( idx );
          wFilename.setText( string[ 0 ] );
          wFilemask.setText( string[ 1 ] );
          wExcludeFilemask.setText( string[ 2 ] );
          wFilenameList.remove( idx );
        }
        wFilenameList.removeEmptyRows();
        wFilenameList.setRowNums();
        input.setChanged();
      }
    } );

    // Show the files that are selected at this time...
    wbShowFiles.addSelectionListener( new SelectionAdapter() {
      public void widgetSelected( SelectionEvent e ) {
        ReadExcelHeaderMeta tfii = new ReadExcelHeaderMeta();
        getInfo( tfii );
        String[] files = tfii.getFilePaths( pipelineMeta );
        if ( files != null && files.length > 0 ) {
          EnterSelectionDialog esd = new EnterSelectionDialog( shell, files, "Files read", "Files read:" );
          esd.setViewOnly();
          esd.open();
        } else {
          MessageBox mb = new MessageBox( shell, SWT.OK | SWT.ICON_ERROR );
          mb.setMessage( BaseMessages.getString( PKG, "ReadExcelHeaderDialog.NoFilesFound.DialogMessage" ) );
          mb.setText( BaseMessages.getString( PKG, "System.Dialog.Error.Title" ) );
          mb.open();
        }
      }
    } );

    // Listen to the Browse... button
    wbbFilename.addListener( SWT.Selection, e-> {
        if ( !Utils.isEmpty( wFilemask.getText() ) || !Utils.isEmpty( wExcludeFilemask.getText() ) ) {
          BaseDialog.presentDirectoryDialog( shell, wFilename, pipelineMeta );
        } else {
          BaseDialog.presentFileDialog( shell, wFilename, pipelineMeta,
            new String[] { "*.xlsx", "*.xls"},
            new String[] { BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FileType.Excelfiles" )},
            true
          );
        }
      } );

    // Detect X or ALT-F4 or something that kills this window...
    shell.addShellListener( new ShellAdapter() {
      public void shellClosed( ShellEvent e ) {
        cancel();
      }
    } );

    wContentFolder.setSelection( 0 );

    // Set the shell size, based upon previous time...
    setFileField();
    getData( input );
    ActiveFileField();
    ActiveStartRowField();
    setSize();
    input.setChanged( changed );

    shell.open();
    while ( !shell.isDisposed() ) {
      if ( !display.readAndDispatch() ) {
        display.sleep();
      }
    }
    return transformName;
  }

  private void ActiveStartRowField() {
		wlStartRowField.setEnabled(wStartRowField.getSelection());
		wStartRowSelField.setEnabled(wStartRowField.getSelection());

		wLabelStepStartRow.setEnabled(!wStartRowField.getSelection());
		wTextStartRow.setEnabled(!wStartRowField.getSelection());
  }
  
  private void setStartRowField() {
		try {

      String startRow = wStartRowSelField.getText();

			wStartRowSelField.removeAll();

			IRowMeta r = pipelineMeta.getPrevTransformFields( transformName );
			if (r != null) {
        wStartRowSelField.setItems( r.getFieldNames() );
      }
      
      if ( startRow != null ) {
        wStartRowSelField.setText( startRow );
      }

		} catch (HopException ke) {
			new ErrorDialog(shell, BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FailedToGetFields.DialogTitle"),
					BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FailedToGetFields.DialogMessage"), ke);
		}
	}

  private void setFileField() {
    try {
      if ( !getpreviousFields ) {
        getpreviousFields = true;
        String filename = wFilenameField.getText();
        String wildcard = wWildcardField.getText();
        String excludewildcard = wExcludeWildcardField.getText();

        wFilenameField.removeAll();
        wWildcardField.removeAll();
        wExcludeWildcardField.removeAll();

        IRowMeta r = pipelineMeta.getPrevTransformFields( transformName );
        if ( r != null ) {
          wFilenameField.setItems( r.getFieldNames() );
          wWildcardField.setItems( r.getFieldNames() );
          wExcludeWildcardField.setItems( r.getFieldNames() );
        }
        if ( filename != null ) {
          wFilenameField.setText( filename );
        }
        if ( wildcard != null ) {
          wWildcardField.setText( wildcard );
        }
        if ( excludewildcard != null ) {
          wExcludeWildcardField.setText( excludewildcard );
        }
      }
    } catch ( HopException ke ) {
      new ErrorDialog(
        shell, BaseMessages.getString( PKG, "ReadExcelHeaderDialog.FailedToGetFields.DialogTitle" ), BaseMessages
        .getString( PKG, "ReadExcelHeaderDialog.FailedToGetFields.DialogMessage" ), ke );
    }
  }

  private void ActiveFileField() {
    wlFilenameField.setEnabled( wFileField.getSelection() );
    wFilenameField.setEnabled( wFileField.getSelection() );
    wlWildcardField.setEnabled( wFileField.getSelection() );
    wWildcardField.setEnabled( wFileField.getSelection() );
    wlExcludeWildcardField.setEnabled( wFileField.getSelection() );
    wExcludeWildcardField.setEnabled( wFileField.getSelection() );
    wlFilename.setEnabled( !wFileField.getSelection() );
    wbbFilename.setEnabled( !wFileField.getSelection() );
    wbaFilename.setEnabled( !wFileField.getSelection() );
    wFilename.setEnabled( !wFileField.getSelection() );
    wlFilemask.setEnabled( !wFileField.getSelection() );
    wFilemask.setEnabled( !wFileField.getSelection() );
    wlExcludeFilemask.setEnabled( !wFileField.getSelection() );
    wExcludeFilemask.setEnabled( !wFileField.getSelection() );
    wlFilenameList.setEnabled( !wFileField.getSelection() );
    wbdFilename.setEnabled( !wFileField.getSelection() );
    wbeFilename.setEnabled( !wFileField.getSelection() );
    wbShowFiles.setEnabled( !wFileField.getSelection() );
    wlFilenameList.setEnabled( !wFileField.getSelection() );
    wFilenameList.setEnabled( !wFileField.getSelection() );
    wPreview.setEnabled( !wFileField.getSelection() );
    wlIncludeSubFolder.setEnabled( wFileField.getSelection() );
    wIncludeSubFolder.setEnabled( wFileField.getSelection() );

  }

  /**
   * Read the data from the ReadExcelHeaderMeta object and show it in this dialog.
   *
   * @param meta The TextFileInputMeta object to obtain the data from.
   */
  public void getData( ReadExcelHeaderMeta meta ) {
    final ReadExcelHeaderMeta in = meta;

    if ( in.getFileName() != null ) {
      wFilenameList.removeAll();

      for ( int i = 0; i < meta.getFileName().length; i++ ) {
        wFilenameList.add( new String[] {
          in.getFileName()[ i ], in.getFileMask()[ i ], in.getExcludeFileMask()[ i ],
          in.getRequiredFilesDesc( in.getFileRequired()[ i ] ),
          in.getRequiredFilesDesc( in.getIncludeSubFolders()[ i ] ) } );
      }

      wdoNotFailIfNoFile.setSelection( in.isdoNotFailIfNoFile() );
      wFilenameList.removeEmptyRows();
      wFilenameList.setRowNums();
      wFilenameList.optWidth( true );

      wFileField.setSelection( in.isFileField() );
      

      if ( in.getDynamicFilenameField() != null ) {
        wFilenameField.setText( in.getDynamicFilenameField() );
      }
      if ( in.getDynamicWildcardField() != null ) {
        wWildcardField.setText( in.getDynamicWildcardField() );
      }
      if ( in.getDynamicExcludeWildcardField() != null ) {
        wExcludeWildcardField.setText( in.getDynamicExcludeWildcardField() );
      }
      wIncludeSubFolder.setSelection( in.isDynamicIncludeSubFolders() );


    }
    if (meta.isStartRowField()) {
      wStartRowSelField.setText(meta.getStartRowFieldName());
    } else {
      wTextStartRow.setText(String.valueOf(meta.getStartRow()));
    }
    wStartRowField.setSelection(meta.isStartRowField());

    wTextSampleRows.setText(String.valueOf(meta.getSampleRows()));

    wTransformName.selectAll();
    wTransformName.setFocus();
  }

  private void cancel() {
    transformName = null;
    input.setChanged( changed );
    dispose();
  }

  private void ok() {
    if ( Utils.isEmpty( wTransformName.getText() ) ) {
      return;
    }

    getInfo( input );
    dispose();
  }

  private void getInfo( ReadExcelHeaderMeta in ) {
    transformName = wTransformName.getText(); // return value

    int nrfiles = wFilenameList.getItemCount();
    in.allocate( nrfiles );

    in.setFileName( wFilenameList.getItems( 0 ) );
    in.setFileMask( wFilenameList.getItems( 1 ) );
    in.setExcludeFileMask( wFilenameList.getItems( 2 ) );
    in.setFileRequired( wFilenameList.getItems( 3 ) );
    in.setIncludeSubFolders( wFilenameList.getItems( 4 ) );

    in.setSampleRows(Integer.parseInt(wTextSampleRows.getText()));

    in.setStartRowField(wStartRowField.getSelection());
    if (in.isStartRowField()) {
      in.setStartRowFieldName(wStartRowSelField.getText());
    } else {
      in.setStartRow(Integer.parseInt(wTextStartRow.getText()));
    }

    in.setDynamicFilenameField( wFilenameField.getText() );
    in.setDynamicWildcardField( wWildcardField.getText() );
    in.setDynamicExcludeWildcardField( wExcludeWildcardField.getText() );
    in.setFileField( wFileField.getSelection() );

    in.setDynamicIncludeSubFolders( wIncludeSubFolder.getSelection() );
    in.setdoNotFailIfNoFile( wdoNotFailIfNoFile.getSelection() );
  }

  // Preview the data
  private void preview() {
    // Create the XML input transform
    ReadExcelHeaderMeta oneMeta = new ReadExcelHeaderMeta();
    getInfo( oneMeta );

    PipelineMeta previewMeta = PipelinePreviewFactory.generatePreviewPipeline( pipelineMeta, pipelineMeta.getMetadataProvider(),
      oneMeta, wTransformName.getText() );

    EnterNumberDialog numberDialog =
      new EnterNumberDialog( shell, props.getDefaultPreviewSize(),
        BaseMessages.getString( PKG, "ReadExcelHeaderDialog.PreviewSize.DialogTitle" ),
        BaseMessages.getString( PKG, "ReadExcelHeaderDialog.PreviewSize.DialogMessage" ) );
    int previewSize = numberDialog.open();
    if ( previewSize > 0 ) {
      PipelinePreviewProgressDialog progressDialog =
        new PipelinePreviewProgressDialog(
          shell, previewMeta, new String[] { wTransformName.getText() }, new int[] { previewSize } );
      progressDialog.open();

      if ( !progressDialog.isCancelled() ) {
        Pipeline pipeline = progressDialog.getPipeline();
        String loggingText = progressDialog.getLoggingText();

        if ( pipeline.getResult() != null && pipeline.getResult().getNrErrors() > 0 ) {
          EnterTextDialog etd =
            new EnterTextDialog( shell, BaseMessages.getString( PKG, "System.Dialog.Error.Title" ), BaseMessages
              .getString( PKG, "ReadExcelHeaderDialog.ErrorInPreview.DialogMessage" ), loggingText, true );
          etd.setReadOnly();
          etd.open();
        }

        PreviewRowsDialog prd =
          new PreviewRowsDialog(
            shell, pipelineMeta, SWT.NONE, wTransformName.getText(), progressDialog.getPreviewRowsMeta( wTransformName
            .getText() ), progressDialog.getPreviewRows( wTransformName.getText() ), loggingText );
        prd.open();
      }
    }
  }
}