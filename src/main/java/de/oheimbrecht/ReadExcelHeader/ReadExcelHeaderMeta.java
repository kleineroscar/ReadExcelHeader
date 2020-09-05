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

import org.apache.commons.vfs2.FileObject;
import org.apache.hop.core.CheckResult;
import org.apache.hop.core.ICheckResult;
import org.apache.hop.core.Const;
import org.apache.hop.core.annotations.Transform;
import org.apache.hop.core.exception.HopException;
import org.apache.hop.core.exception.HopTransformException;
import org.apache.hop.core.exception.HopXmlException;
import org.apache.hop.core.fileinput.FileInputList;
import org.apache.hop.core.row.IRowMeta;
import org.apache.hop.core.row.IValueMeta;
import org.apache.hop.core.row.value.ValueMetaString;
import org.apache.hop.core.util.Utils;
import org.apache.hop.core.variables.IVariables;
import org.apache.hop.core.vfs.HopVfs;
import org.apache.hop.core.xml.XmlHandler;
import org.apache.hop.i18n.BaseMessages;
import org.apache.hop.metadata.api.IHopMetadataProvider;
import org.apache.hop.pipeline.Pipeline;
import org.apache.hop.pipeline.PipelineMeta;
import org.apache.hop.resource.ResourceDefinition;
import org.apache.hop.resource.ResourceEntry;
import org.apache.hop.resource.ResourceEntry.ResourceType;
import org.apache.hop.resource.IResourceNaming;
import org.apache.hop.resource.ResourceReference;
import org.apache.hop.pipeline.transform.*;
import org.w3c.dom.Node;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Transform(
        id = "ReadExcelHeader",
        image = "REH.svg",
        i18nPackageName = "i18n:de.oheimbrecht.ReadExcelHeader",
        name = "BaseTransform.TypeLongDesc.ReadExcelHeader",
        description = "BaseTransform.TypeTooltipDesc.ReadExcelHeader",
        categoryDescription = "i18n:org.apache.hop.pipeline.transform:BaseTransform.Category.Input",
        documentationUrl = "https://github.com/kleineroscar/ReadExcelHeader"
)
public class ReadExcelHeaderMeta extends BaseTransformMeta implements ITransformMeta<ReadExcelHeader, ReadExcelHeaderData> {
  private static Class<?> PKG = ReadExcelHeaderMeta.class; // for i18n purposes, needed by Translator!!

  public static final String[] RequiredFilesDesc = new String[] {
    BaseMessages.getString( PKG, "System.Combo.No" ), BaseMessages.getString( PKG, "System.Combo.Yes" ) };
  public static final String[] RequiredFilesCode = new String[] { "N", "Y" };

  private static final String NO = "N";

  private static final String YES = "Y";

  /**
   * Array of filenames
   */
  private String[] fileName;

  /**
   * Wildcard or filemask (regular expression)
   */
  private String[] fileMask;

  /**
   * Wildcard or filemask to exclude (regular expression)
   */
  private String[] excludeFileMask;

  /**
   * Array of boolean values as string, indicating if a file is required.
   */
  private String[] fileRequired;

  /**
   * Array of boolean values as string, indicating if we need to fetch sub folders.
   */
  private String[] includeSubFolders;

  /**
   * The name of the field in the output containing the filename
   */
  private String filenameField;

  /**
   * Flag indicating that a field should be used for the start row
   */
  private boolean startrowfield;

  /**
   * The name of the field in the output containing the start row
   */
  private String startRowFieldName;

  /**
   * The start row value
   */
  private int startRow;
  

  /**
   * How many rows to sample
   */
  private int sampleRows;

  private String dynamicFilenameField;

  private String dynamicWildcardField;
  private String dynamicExcludeWildcardField;

  /**
   * file name from previous fields
   **/
  private boolean filefield;

  private boolean dynamicIncludeSubFolders;

  /**
   * Flag : do not fail if no file
   */
  private boolean doNotFailIfNoFile;

  public ReadExcelHeaderMeta() {
    super(); // allocate BaseTransformMeta
  }

  /**
   * @return the doNotFailIfNoFile flag
   */
  public boolean isdoNotFailIfNoFile() {
    return doNotFailIfNoFile;
  }

  /**
   * @param doNotFailIfNoFile the doNotFailIfNoFile to set
   */
  public void setdoNotFailIfNoFile( boolean doNotFailIfNoFile ) {
    this.doNotFailIfNoFile = doNotFailIfNoFile;
  }

  /**
   * @return Returns the filenameField.
   */
  public String getFilenameField() {
    return filenameField;
  }

  /**
   * @param dynamicFilenameField The dynamic filename field to set.
   */
  public void setDynamicFilenameField( String dynamicFilenameField ) {
    this.dynamicFilenameField = dynamicFilenameField;
  }

  /**
   * @param dynamicWildcardField The dynamic wildcard field to set.
   */
  public void setDynamicWildcardField( String dynamicWildcardField ) {
    this.dynamicWildcardField = dynamicWildcardField;
  }

  /**
   * @return Returns the dynamic filename field (from previous transforms)
   */
  public String getDynamicFilenameField() {
    return dynamicFilenameField;
  }

  /**
   * @return Returns the dynamic wildcard field (from previous transforms)
   */
  public String getDynamicWildcardField() {
    return dynamicWildcardField;
  }

  public String getDynamicExcludeWildcardField() {
    return this.dynamicExcludeWildcardField;
  }

  /**
   * @param dynamicExcludeWildcardField The dynamic excludeWildcard field to set.
   */
  public void setDynamicExcludeWildcardField( String dynamicExcludeWildcardField ) {
    this.dynamicExcludeWildcardField = dynamicExcludeWildcardField;
  }


  /**
   * @return Returns the File field.
   */
  public boolean isFileField() {
    return filefield;
  }

  /**
   * @param filefield The filefield to set.
   */
  public void setFileField( boolean filefield ) {
    this.filefield = filefield;
  }

  public boolean isDynamicIncludeSubFolders() {
    return dynamicIncludeSubFolders;
  }

  public void setDynamicIncludeSubFolders( boolean dynamicIncludeSubFolders ) {
    this.dynamicIncludeSubFolders = dynamicIncludeSubFolders;
  }

  /**
   * @return Returns the fileMask.
   */
  public String[] getFileMask() {
    return fileMask;
  }

  /**
   * @return Returns the fileRequired.
   */
  public String[] getFileRequired() {
    return fileRequired;
  }

  /**
   * @param fileMask The fileMask to set.
   */
  public void setFileMask( String[] fileMask ) {
    this.fileMask = fileMask;
  }

  /**
   * @param excludeFileMask The excludeFileMask to set.
   */
  public void setExcludeFileMask( String[] excludeFileMask ) {
    this.excludeFileMask = excludeFileMask;
  }


  /**
   * @return Returns the excludeFileMask.
   */
  public String[] getExcludeFileMask() {
    return excludeFileMask;
  }

  /**
   * @param fileRequiredin The fileRequired to set.
   */
  public void setFileRequired( String[] fileRequiredin ) {
    this.fileRequired = new String[ fileRequiredin.length ];
    for ( int i = 0; i < fileRequiredin.length; i++ ) {
      this.fileRequired[ i ] = getRequiredFilesCode( fileRequiredin[ i ] );
    }
  }

  /**
   * @return Returns the fileName.
   */
  public String[] getFileName() {
    return fileName;
  }

  /**
   * @param fileName The fileName to set.
   */
  public void setFileName( String[] fileName ) {
    this.fileName = fileName;
  }

  /**
   * @return Returns if a field is used for startRow.
   */
  public boolean isStartRowField() {
      return startrowfield;
  }

  /**
   * @param startrowfield The startrowfield to set.
   */
  public void setStartRowField( boolean startrowfield ) {
      this.startrowfield = startrowfield;
  }

  /**
   * @return Returns the start row field name.
   */
  public String getStartRowFieldName() {
      return startRowFieldName;
  }

  /**
   * @param startRowFieldName The startRowFieldName to set.
   */
  public void setStartRowFieldName( String startRowFieldName ) {
      this.startRowFieldName = startRowFieldName;
  }

  /**
   * @return Returns the start row if the entry is manual.
   */
  public int getStartRow() {
      return startRow;
  }

  /**
   * @param startRow The startRow to set.
   */
  public void setStartRow(int startRow) {
      this.startRow = startRow;
  }  

  /**
   * @return Returns the number of sample rows.
   */
  public int getSampleRows() {
      return sampleRows;
  }

  /**
   * @param sampleRows The fileName to set.
   */
  public void setSampleRows(int sampleRows) {
      this.sampleRows = sampleRows;
  }

  public String getRequiredFilesDesc( String tt ) {
    if ( tt == null ) {
      return RequiredFilesDesc[ 0 ];
    }
    if ( tt.equals( RequiredFilesCode[ 1 ] ) ) {
      return RequiredFilesDesc[ 1 ];
    } else {
      return RequiredFilesDesc[ 0 ];
    }
  }

  public void setIncludeSubFolders( String[] includeSubFoldersin ) {
    this.includeSubFolders = new String[ includeSubFoldersin.length ];
    for ( int i = 0; i < includeSubFoldersin.length; i++ ) {
      this.includeSubFolders[ i ] = getRequiredFilesCode( includeSubFoldersin[ i ] );
    }
  }

  public String getRequiredFilesCode( String tt ) {
    if ( tt == null ) {
      return RequiredFilesCode[ 0 ];
    }
    if ( tt.equals( RequiredFilesDesc[ 1 ] ) ) {
      return RequiredFilesCode[ 1 ];
    } else {
      return RequiredFilesCode[ 0 ];
    }
  }

  @Override
  public void loadXml( Node transformNode, IHopMetadataProvider metadataProvider ) throws HopXmlException {
    readData( transformNode );
  }

  @Override
  public Object clone() {
    ReadExcelHeaderMeta retval = (ReadExcelHeaderMeta) super.clone();

    int nrfiles = fileName.length;

    retval.allocate( nrfiles );

    System.arraycopy( fileName, 0, retval.fileName, 0, nrfiles );
    System.arraycopy( fileMask, 0, retval.fileMask, 0, nrfiles );
    System.arraycopy( excludeFileMask, 0, retval.excludeFileMask, 0, nrfiles );
    System.arraycopy( fileRequired, 0, retval.fileRequired, 0, nrfiles );
    System.arraycopy( includeSubFolders, 0, retval.includeSubFolders, 0, nrfiles );

    return retval;
  }

  public void allocate( int nrfiles ) {
    fileName = new String[ nrfiles ];
    fileMask = new String[ nrfiles ];
    excludeFileMask = new String[ nrfiles ];
    fileRequired = new String[ nrfiles ];
    includeSubFolders = new String[ nrfiles ];
  }

  @Override
  public void setDefault() {
    int nrfiles = 0;
    doNotFailIfNoFile = false;
    filefield = false;
    dynamicFilenameField = "";
    dynamicWildcardField = "";
    dynamicIncludeSubFolders = false;
    dynamicExcludeWildcardField = "";

    sampleRows = 10;
    startrowfield = false;
    startRowFieldName = "";
    startRow = 0;

    allocate( nrfiles );

    for ( int i = 0; i < nrfiles; i++ ) {
      fileName[ i ] = "filename" + ( i + 1 );
      fileMask[ i ] = "";
      excludeFileMask[ i ] = "";
      fileRequired[ i ] = NO;
      includeSubFolders[ i ] = NO;
    }
  }

  @Override
  public void getFields( IRowMeta row, String name, IRowMeta[] info, TransformMeta nextTransform,
                         IVariables variables, IHopMetadataProvider metadataProvider ) throws HopTransformException {

    // the workbookName
    IValueMeta workbookName = new ValueMetaString( "workbookName" );
    workbookName.setLength( 500 );
    workbookName.setPrecision( -1 );
    workbookName.setOrigin( name );
    row.addValueMeta( workbookName );

    // the sheetName
    IValueMeta sheetName = new ValueMetaString( "sheetName" );
    sheetName.setLength( 500 );
    sheetName.setPrecision( -1 );
    sheetName.setOrigin( name );
    row.addValueMeta( sheetName );

    // the columnName
    IValueMeta columnName = new ValueMetaString( "columnName" );
    columnName.setLength( 500 );
    columnName.setPrecision( -1 );
    columnName.setOrigin( name );
    row.addValueMeta( columnName );

    // the columnType
    IValueMeta columnType = new ValueMetaString( "columnType" );
    columnType.setLength( 500 );
    columnType.setPrecision( -1 );
    columnType.setOrigin( name );
    row.addValueMeta( columnType );

    // the columnDataFormat
    IValueMeta columnDataFormat = new ValueMetaString( "columnDataFormat" );
    columnDataFormat.setLength( 500 );
    columnDataFormat.setPrecision( -1 );
    columnDataFormat.setOrigin( name );
    row.addValueMeta( columnDataFormat );
  }

  @Override
  public String getXml() {
    StringBuilder retval = new StringBuilder( 300 );

    retval.append( "    " ).append( XmlHandler.addTagValue( "doNotFailIfNoFile", doNotFailIfNoFile ) );
    retval.append( "    " ).append( XmlHandler.addTagValue( "filefield", filefield ) );
    retval.append( "    " ).append( XmlHandler.addTagValue( "filename_Field", dynamicFilenameField ) );
    retval.append( "    " ).append( XmlHandler.addTagValue( "wildcard_Field", dynamicWildcardField ) );
    retval
      .append( "    " ).append( XmlHandler.addTagValue( "exclude_wildcard_Field", dynamicExcludeWildcardField ) );
    retval.append( "    " ).append(
      XmlHandler.addTagValue( "dynamic_include_subfolders", dynamicIncludeSubFolders ) );
    retval.append( "    <file>" ).append( Const.CR );

    for ( int i = 0; i < fileName.length; i++ ) {
      retval.append( "      " ).append( XmlHandler.addTagValue( "name", fileName[ i ] ) );
      retval.append( "      " ).append( XmlHandler.addTagValue( "filemask", fileMask[ i ] ) );
      retval.append( "      " ).append( XmlHandler.addTagValue( "exclude_filemask", excludeFileMask[ i ] ) );
      retval.append( "      " ).append( XmlHandler.addTagValue( "file_required", fileRequired[ i ] ) );
      retval.append( "      " ).append( XmlHandler.addTagValue( "include_subfolders", includeSubFolders[ i ] ) );
    }
    retval.append( "    </file>" ).append( Const.CR );

    retval.append( "    " ).append( XmlHandler.addTagValue("startrowfield", startrowfield) );
    retval.append( "    " ).append( XmlHandler.addTagValue("startrowfieldname", startRowFieldName) );
    retval.append( "    " ).append( XmlHandler.addTagValue( "startrow", startRow ) );
    retval.append( "    " ).append( XmlHandler.addTagValue( "sampleRows", sampleRows ) );
    return retval.toString();
  }

  private void readData( Node transformNode ) throws HopXmlException {
    try {
      doNotFailIfNoFile = "Y".equalsIgnoreCase( XmlHandler.getTagValue( transformNode, "doNotFailIfNoFile" ) );
      filefield = "Y".equalsIgnoreCase( XmlHandler.getTagValue( transformNode, "filefield" ) );
      dynamicFilenameField = XmlHandler.getTagValue( transformNode, "filename_Field" );
      dynamicWildcardField = XmlHandler.getTagValue( transformNode, "wildcard_Field" );
      dynamicExcludeWildcardField = XmlHandler.getTagValue( transformNode, "exclude_wildcard_Field" );
      dynamicIncludeSubFolders =
        "Y".equalsIgnoreCase( XmlHandler.getTagValue( transformNode, "dynamic_include_subfolders" ) );

      Node filenode = XmlHandler.getSubNode( transformNode, "file" );
      int nrfiles = XmlHandler.countNodes( filenode, "name" );

      allocate( nrfiles );

      for ( int i = 0; i < nrfiles; i++ ) {
        Node filenamenode = XmlHandler.getSubNodeByNr( filenode, "name", i );
        Node filemasknode = XmlHandler.getSubNodeByNr( filenode, "filemask", i );
        Node excludefilemasknode = XmlHandler.getSubNodeByNr( filenode, "exclude_filemask", i );
        Node fileRequirednode = XmlHandler.getSubNodeByNr( filenode, "file_required", i );
        Node includeSubFoldersnode = XmlHandler.getSubNodeByNr( filenode, "include_subfolders", i );
        fileName[ i ] = XmlHandler.getNodeValue( filenamenode );
        fileMask[ i ] = XmlHandler.getNodeValue( filemasknode );
        excludeFileMask[ i ] = XmlHandler.getNodeValue( excludefilemasknode );
        fileRequired[ i ] = XmlHandler.getNodeValue( fileRequirednode );
        includeSubFolders[ i ] = XmlHandler.getNodeValue( includeSubFoldersnode );
      }

      startrowfield = "Y".equalsIgnoreCase(XmlHandler.getTagValue(filenode, "startrowfield"));

      if (isStartRowField()) {
        startRowFieldName = XmlHandler.getTagValue(filenode, "startrowfieldname");
      } else {
        startRow = Integer.parseInt(XmlHandler.getTagValue(filenode, "startrow"));
      }
      sampleRows = Integer.parseInt(XmlHandler.getTagValue(filenode, "sampleRows"));
    } catch ( Exception e ) {
      throw new HopXmlException( "Unable to load transform info from XML", e );
    }
  }

  private boolean[] includeSubFolderBoolean() {
    int len = fileName.length;
    boolean[] includeSubFolderBoolean = new boolean[ len ];
    for ( int i = 0; i < len; i++ ) {
      includeSubFolderBoolean[ i ] = YES.equalsIgnoreCase( includeSubFolders[ i ] );
    }
    return includeSubFolderBoolean;
  }

  public String[] getIncludeSubFolders() {
    return includeSubFolders;
  }

  public String[] getFilePaths( IVariables variables ) {
    return FileInputList.createFilePathList(
      variables, fileName, fileMask, excludeFileMask, fileRequired, includeSubFolderBoolean() );
  }

  public FileInputList getFileList( IVariables variables ) {
    return FileInputList.createFileList(
      variables, fileName, fileMask, excludeFileMask, fileRequired, includeSubFolderBoolean() );
  }

  public FileInputList getDynamicFileList( IVariables variables, String[] filename, String[] filemask,
                                           String[] excludefilemask, String[] filerequired, boolean[] includesubfolders ) {
    return FileInputList.createFileList(
      variables, filename, filemask, excludefilemask, filerequired, includesubfolders );
  }

  @Override
  public void check( List<ICheckResult> remarks, PipelineMeta pipelineMeta, TransformMeta transformMeta,
                     IRowMeta prev, String[] input, String[] output, IRowMeta info, IVariables variables,
                     IHopMetadataProvider metadataProvider ) {
    CheckResult cr;

    // See if we get input...
    if ( filefield ) {
      if ( input.length > 0 ) {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_OK, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.InputOk" ), transformMeta );
      } else {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_ERROR, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.InputErrorKo" ), transformMeta );
      }
      remarks.add( cr );

      if ( Utils.isEmpty( dynamicFilenameField ) ) {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_ERROR, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.FolderFieldnameMissing" ), transformMeta );
      } else {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_OK, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.FolderFieldnameOk" ), transformMeta );
      }
      remarks.add( cr );

    } else {

      if ( input.length > 0 ) {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_ERROR, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.NoInputError" ), transformMeta );
      } else {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_OK, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.NoInputOk" ), transformMeta );
      }

      remarks.add( cr );

      // check specified file names
      FileInputList fileList = getFileList( pipelineMeta );
      if ( fileList.nrOfFiles() == 0 ) {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_ERROR, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.ExpectedFilesError" ), transformMeta );
      } else {
        cr =
          new CheckResult( ICheckResult.TYPE_RESULT_OK, BaseMessages.getString(
            PKG, "ReadExcelHeaderMeta.CheckResult.ExpectedFilesOk", "" + fileList.nrOfFiles() ), transformMeta );
      }
      remarks.add( cr );
    }

    if ( startrowfield) {
        if (!startRowFieldName.isEmpty()) {
            cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received a field with a row to start on.", transformMeta);
            remarks.add(cr);
        } else {
            cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No start row field received!", transformMeta);
            remarks.add(cr);
        }
    } else {
        if (startRow != -1) {
            cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received a row to start on.", transformMeta);
            remarks.add(cr);
        } else {
            cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No start row received!", transformMeta);
            remarks.add(cr);
        }
    }

    if (sampleRows != -1) {
        cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received number of rows to sample on.", transformMeta);
        remarks.add(cr);
    } else {
        cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No number of sample rows received!", transformMeta);
        remarks.add(cr);
    }
  }

  @Override
  public List<ResourceReference> getResourceDependencies( PipelineMeta pipelineMeta, TransformMeta transformInfo ) {
    List<ResourceReference> references = new ArrayList<ResourceReference>( 5 );
    ResourceReference reference = new ResourceReference( transformInfo );
    references.add( reference );

    String[] files = getFilePaths( pipelineMeta );
    if ( files != null ) {
      for ( int i = 0; i < files.length; i++ ) {
        reference.getEntries().add( new ResourceEntry( files[ i ], ResourceType.FILE ) );
      }
    }
    return references;
  }

  @Override
  public ReadExcelHeader createTransform( TransformMeta transformMeta, ReadExcelHeaderData data, int cnr,
                                       PipelineMeta pipelineMeta, Pipeline pipeline ) {
    return new ReadExcelHeader( transformMeta, this, data, cnr, pipelineMeta, pipeline );
  }

  @Override
  public ReadExcelHeaderData getTransformData() {
    return new ReadExcelHeaderData();
  }

  /**
   * @param variables                   the variable space to use
   * @param definitions
   * @param iResourceNaming
   * @param metadataProvider               the metadataProvider in which non-hop metadata could reside.
   * @return the filename of the exported resource
   */
  @Override
  public String exportResources( IVariables variables, Map<String, ResourceDefinition> definitions,
                                 IResourceNaming iResourceNaming, IHopMetadataProvider metadataProvider ) throws HopException {
    try {
      // The object that we're modifying here is a copy of the original!
      // So let's change the filename from relative to absolute by grabbing the file object...
      // In case the name of the file comes from previous transforms, forget about this!
      //
      if ( !filefield ) {

        // Replace the filename ONLY (folder or filename)
        //
        for ( int i = 0; i < fileName.length; i++ ) {
          FileObject fileObject = HopVfs.getFileObject( variables.environmentSubstitute( fileName[ i ] ) );
          fileName[ i ] = iResourceNaming.nameResource( fileObject, variables, Utils.isEmpty( fileMask[ i ] ) );
        }
      }
      return null;
    } catch ( Exception e ) {
      throw new HopException( e );
    }
  }
}
