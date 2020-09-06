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
import org.apache.commons.vfs2.provider.local.LocalFile;
import org.apache.hop.core.Const;
import org.apache.hop.core.exception.HopException;
import org.apache.hop.core.exception.HopTransformException;
import org.apache.hop.core.fileinput.FileInputList;
import org.apache.hop.core.row.RowDataUtil;
import org.apache.hop.core.row.RowMeta;
import org.apache.hop.core.util.Utils;
import org.apache.hop.core.vfs.HopVfs;
import org.apache.hop.i18n.BaseMessages;
import org.apache.hop.pipeline.Pipeline;
import org.apache.hop.pipeline.PipelineMeta;
import org.apache.hop.pipeline.transform.BaseTransform;
import org.apache.hop.pipeline.transform.ITransform;
import org.apache.hop.pipeline.transform.TransformMeta;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

// import org.apache.poi.xssf.usermodel.*;

/**
 * Read all sorts of text files, convert them to rows and writes these to one or more output streams.
 *
 * @author Matt
 * @since 4-apr-2003
 */
public class ReadExcelHeader extends BaseTransform<ReadExcelHeaderMeta, ReadExcelHeaderData> implements ITransform<ReadExcelHeaderMeta, ReadExcelHeaderData> {

  private static Class<?> PKG = ReadExcelHeaderMeta.class; // for i18n purposes, needed by Translator!!

  private int startRow = -1;
	private int sampleRows = -1;

  public ReadExcelHeader( TransformMeta transformMeta, ReadExcelHeaderMeta meta, ReadExcelHeaderData data, int copyNr, PipelineMeta pipelineMeta,
                       Pipeline pipeline ) {
    super( transformMeta, meta, data, copyNr, pipelineMeta, pipeline );
  }

  /**
   * Build an empty row based on the meta-data...
   *
   * @return
   */

  private Object[] buildEmptyRow() {
    Object[] rowData = RowDataUtil.allocateRowData( data.outputRowMeta.size() );

    return rowData;
  }

  public boolean processRow() throws HopException {
    if ( !meta.isFileField() ) {
      if ( data.filenr >= data.filessize ) {
        setOutputDone();
        return false;
      }
    } else {
      if ( data.filenr >= data.filessize ) {
        // Grab one row from previous transform ...
        data.readrow = getRow();
      }

      if ( data.readrow == null ) {
        setOutputDone();
        return false;
      }

      if ( first ) {
        first = false;

        data.inputRowMeta = getInputRowMeta();
        data.outputRowMeta = data.inputRowMeta.clone();
        meta.getFields( data.outputRowMeta, getTransformName(), null, null, this, metadataProvider );

        // Get total previous fields
        data.totalpreviousfields = data.inputRowMeta.size();

        // Check is filename field is provided
        if ( Utils.isEmpty( meta.getDynamicFilenameField() ) ) {
          logError( BaseMessages.getString( PKG, "ReadExcelHeader.Log.NoField" ) );
          throw new HopException( BaseMessages.getString( PKG, "ReadExcelHeader.Log.NoField" ) );
        }

        // cache the position of the field
        if ( data.indexOfFilenameField < 0 ) {
          data.indexOfFilenameField = data.inputRowMeta.indexOfValue( meta.getDynamicFilenameField() );
          if ( data.indexOfFilenameField < 0 ) {
            // The field is unreachable !
            logError( BaseMessages.getString( PKG, "ReadExcelHeader.Log.ErrorFindingField", meta
              .getDynamicFilenameField() ) );
            throw new HopException( BaseMessages.getString(
              PKG, "ReadExcelHeader.Exception.CouldnotFindField", meta.getDynamicFilenameField() ) );
          }
        }

        // If wildcard field is specified, Check if field exists
        if ( !Utils.isEmpty( meta.getDynamicWildcardField() ) ) {
          if ( data.indexOfWildcardField < 0 ) {
            data.indexOfWildcardField = data.inputRowMeta.indexOfValue( meta.getDynamicWildcardField() );
            if ( data.indexOfWildcardField < 0 ) {
              // The field is unreachable !
              logError( BaseMessages.getString( PKG, "ReadExcelHeader.Log.ErrorFindingField" )
                + "[" + meta.getDynamicWildcardField() + "]" );
              throw new HopException( BaseMessages.getString(
                PKG, "ReadExcelHeader.Exception.CouldnotFindField", meta.getDynamicWildcardField() ) );
            }
          }
        }
        // If ExcludeWildcard field is specified, Check if field exists
        if ( !Utils.isEmpty( meta.getDynamicExcludeWildcardField() ) ) {
          if ( data.indexOfExcludeWildcardField < 0 ) {
            data.indexOfExcludeWildcardField =
              data.inputRowMeta.indexOfValue( meta.getDynamicExcludeWildcardField() );
            if ( data.indexOfExcludeWildcardField < 0 ) {
              // The field is unreachable !
              logError( BaseMessages.getString( PKG, "ReadExcelHeader.Log.ErrorFindingField" )
                + "[" + meta.getDynamicExcludeWildcardField() + "]" );
              throw new HopException( BaseMessages.getString(
                PKG, "ReadExcelHeader.Exception.CouldnotFindField", meta.getDynamicExcludeWildcardField() ) );
            }
          }
        }
      } // end if first
    } 

    try {
      Object[] inputRow = buildEmptyRow();
      if ( meta.isFileField() ) {
        if ( data.filenr >= data.filessize ) {
          // Get value of dynamic filename field ...
          String filename = getInputRowMeta().getString( data.readrow, data.indexOfFilenameField );
          String wildcard = "";
          if ( data.indexOfWildcardField >= 0 ) {
            wildcard = getInputRowMeta().getString( data.readrow, data.indexOfWildcardField );
          }
          String excludewildcard = "";
          if ( data.indexOfExcludeWildcardField >= 0 ) {
            excludewildcard = getInputRowMeta().getString( data.readrow, data.indexOfExcludeWildcardField );
          }

          String[] filesname = { filename };
          String[] filesmask = { wildcard };
          String[] excludefilesmask = { excludewildcard };
          String[] filesrequired = { "N" };
          boolean[] includesubfolders = { meta.isDynamicIncludeSubFolders() };
          // Get files list
          data.files =
            meta.getDynamicFileList(
              this, filesname, filesmask, excludefilesmask, filesrequired, includesubfolders );
          data.filessize = data.files.nrOfFiles();
          data.filenr = 0;
        }

        // Clone current input row
        inputRow = data.readrow.clone();
      } else {
        if (meta.isStartRowField()) {
					inputRow = data.readrow.clone();
					inputRow = RowDataUtil.resizeArray(inputRow, data.outputRowMeta.size());
				} else {
					inputRow = RowDataUtil.allocateRowData(data.outputRowMeta.size());
				}
      }
      int indexOfStartRowField = -1;
		  try {
		  	if (meta.isStartRowField()) {
		  		indexOfStartRowField = data.inputRowMeta.indexOfValue(meta.getStartRowFieldName());
		  		if (indexOfStartRowField < 0) {
		  			// The field is unreachable !
		  			logError(BaseMessages.getString( PKG, "ReadExcelHeader.Log.ErrorFindingStartRowField",
		  					meta.getStartRowFieldName()));
		  			throw new HopException(BaseMessages.getString( PKG, 
		  					"ReadExcelHeader.Exception.CouldnotFindStartRowField", meta.getStartRowFieldName()));
		  		}
		  		startRow = Integer.parseInt(data.inputRowMeta.getString(inputRow, indexOfStartRowField));
		  	} else {
		  		startRow = Integer.parseInt(environmentSubstitute(String.valueOf(meta.getStartRow())));
		  	}
		  } catch (NumberFormatException nfe) {
		  	logError("StartRow couldn't be parsed");
		  	logDebug("Startrow: " + meta.getStartRow());
		  	logDebug("Startrow from field: " + data.inputRowMeta.getString(inputRow, indexOfStartRowField));
		  	logDebug("Startrow field index: " + indexOfStartRowField);
		  	throw new HopException("StartRow couldn't be parsed.");
		  } catch (Exception e) {
		  	logDebug("Isstartrowfield: " + meta.isStartRowField());
		  	logDebug("startrowfieldname: " + meta.getStartRowFieldName());
		  	logDebug("startrow: " + meta.getStartRow());
		  	logDebug("data.inputRowMeta null?: " + (data.inputRowMeta == null ? "yes" : "no"));
		  	throw new HopException("Some other error while parsing startRow" + e.getMessage());
		  }
      if ( data.filessize > 0 ) {
        
        data.file = data.files.getFile( data.filenr );
        InputStream excelFileInputStream = null;
        XSSFWorkbook workbook1 = null;

        try {
          // HopVfs.getFilename( data.file )

          FileObject fileObject = HopVfs.getFileObject(HopVfs.getFilename(data.file));
          if (fileObject instanceof LocalFile) {
            //This might reduce memory usage
            logDebug("Local file");
            String localFilename = HopVfs.getFilename( data.file );
            File excelFile = new File(localFilename);
            excelFileInputStream = new FileInputStream(excelFile);
            workbook1 = new XSSFWorkbook(excelFileInputStream);
          } else {
            logDebug("VFS file");
            excelFileInputStream = HopVfs.getInputStream(HopVfs.getFilename( data.file ));
            workbook1 = new XSSFWorkbook(excelFileInputStream);
          }
          logDebug("successfully read file");
          
        } catch (IOException e) {
          logDebug("Couldn't get file");
          logDebug(e.getMessage());
          throw new HopException("Couldn't read file provided.");
        }
    
        for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
          XSSFSheet sheet;
          XSSFRow row;

          int outputIndex = 0;
          Object[] extraData = new Object[ data.nrTransformFields ];
    
          try {
            sheet = workbook1.getSheetAt(i);
          } catch (Exception e) {
            log.logError("Unable to read sheet\n" + e.getMessage());
            workbook1.close();
            excelFileInputStream.close();
            throw new HopException("Could not read sheet with number: " + i);
          }
          try {
            row = sheet.getRow(startRow);
            log.logRowlevel("Found a sheet with the corresponding header row (from/to): " + row.getFirstCellNum() + "/"
                + row.getLastCellNum());
          } catch (Exception e) {
            log.logDebug("Unable to read row.\nMaybe the row given is empty.\n Providing empty row.");
            extraData[ outputIndex++ ] = HopVfs.getFilename( data.file );
            log.logRowlevel("Got workbook name: " + HopVfs.getFilename( data.file ));
            extraData[ outputIndex++ ] = workbook1.getSheetName(i);
            log.logRowlevel("Got sheet name: " + workbook1.getSheetName(i));

            // Fill up the rest of the columns with pseudo data
            extraData[ outputIndex++ ] = "NO DATA";
            extraData[ outputIndex++ ] = "NO DATA";
            extraData[ outputIndex++ ] = "NO DATA";
            // put the row to the output row stream
            log.logRowlevel(
                "Created the following row: " + Arrays.toString(extraData));
            // Add row data
            Object[] resultRow = RowDataUtil.addRowData( inputRow, data.totalpreviousfields, extraData );
            // Send row
            putRow( data.outputRowMeta, resultRow );
    
            continue;
          }
          for (short j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
            outputIndex = 0;
            // generate output row, make it correct size
            try {
              log.logRowlevel("Processing the next cell with number: " + j);
              extraData[ outputIndex++ ] = HopVfs.getFilename( data.file );
              log.logRowlevel("Got workbook name: " + HopVfs.getFilename( data.file ));
            } catch (Exception e) {
              log.logDebug( e.getMessage() );
              workbook1.close();
              excelFileInputStream.close();
              throw new HopException("Some error while getting the file string " + HopVfs.getFilename( data.file ));
            }
            try {
              extraData[ outputIndex++ ] = workbook1.getSheetName(i);
              log.logRowlevel("Got sheet name: " + workbook1.getSheetName(i));
            } catch (Exception e) {
              log.logDebug(e.getMessage());
              workbook1.close();
              excelFileInputStream.close();
              throw new HopException("Some error while getting the values. With Sheetnumber:" + String.valueOf(i));
            }
            try{
              // Split this into two statements so outputIndex does not incremented imediately
              String tempString = row.getCell(j).toString();
              extraData[ outputIndex++ ] = tempString;
              log.logRowlevel("Got cell header: " + row.getCell(j).toString());
            } catch (Exception e) {
              log.logDebug("Some error while getting the values. With Sheetnumber:" + String.valueOf(i)
                  + " and Column number:" + String.valueOf(j));
              
              // Dont throw an exception here.
              // This is most likely just an empty cell.
              // throw new HopException(e.getMessage());
              extraData[ outputIndex++ ] = "NO DATA";
            }
    
    
            Map<String, String[]> cellInfo = new HashMap<>();
            for (int k = startRow + 1; k <= sampleRows; k++) {
              try {
                XSSFCell cell = sheet.getRow(k).getCell(row.getCell(j).getColumnIndex());
                log.logRowlevel("Adding type and style to list: '" + cell.getCellTypeEnum().toString() + "/"
                    + cell.getCellStyle().getDataFormatString() + "'");
                cellInfo.put(cell.getCellTypeEnum().toString() + cell.getCellStyle().getDataFormatString(),
                    new String[] { cell.getCellTypeEnum().toString(),
                        cell.getCellStyle().getDataFormatString() });
              } catch (Exception e) {
                log.logRowlevel("Couldn't get Field info in row " + String.valueOf(k));
              }
            }
            if (cellInfo.size() == 0) {
              extraData[ outputIndex++ ] = "NO DATA";
              extraData[ outputIndex++ ] = "NO DATA";
            } else if (cellInfo.size() > 1) {
              extraData[ outputIndex++ ] = "STRING";
              extraData[ outputIndex++ ] = "Mixed";
            } else {
              String[] info = (String[]) cellInfo.values().toArray()[0];
              extraData[ outputIndex++ ] = info[0];
              extraData[ outputIndex++ ] = info[1];
            }
    
            // put the row to the output row stream
            log.logRowlevel(
                "Created the following row: " + Arrays.toString(extraData) + " ; number of different data formats=" + cellInfo.size());
            // Add row data
            Object[] resultRow = RowDataUtil.addRowData( inputRow, data.totalpreviousfields, extraData );
            // Send row
            putRow( data.outputRowMeta, resultRow );
          }
        }

        try {
          workbook1.close();
        } catch (Exception e) {
          new HopException("Could not dispose workbook.\n" + e.getMessage());
        }
    
        try {
          excelFileInputStream.close();
        } catch (IOException e) {
          new HopException("Could not dispose FileInputStream.\n" + e.getMessage());
        }

      }
    } catch ( Exception e ) {
      throw new HopTransformException( e );
    }

    data.filenr++;

    if ( checkFeedback( getLinesInput() ) ) {
      if ( log.isBasic() ) {
        logBasic( BaseMessages.getString( PKG, "ReadExcelHeader.Log.NrLine", "" + getLinesInput() ) );
      }
    }

    return true;
  }

  private void handleMissingFiles() throws HopException {
    if ( meta.isdoNotFailIfNoFile() && data.files.nrOfFiles() == 0 ) {
      logBasic( BaseMessages.getString( PKG, "ReadExcelHeader.Log.NoFile" ) );
      return;
    }
    List<FileObject> nonExistantFiles = data.files.getNonExistantFiles();

    if ( nonExistantFiles.size() != 0 ) {
      String message = FileInputList.getRequiredFilesDescription( nonExistantFiles );
      logBasic( "ERROR: Missing " + message );
      throw new HopException( "Following required files are missing: " + message );
    }

    List<FileObject> nonAccessibleFiles = data.files.getNonAccessibleFiles();
    if ( nonAccessibleFiles.size() != 0 ) {
      String message = FileInputList.getRequiredFilesDescription( nonAccessibleFiles );
      logBasic( "WARNING: Not accessible " + message );
      throw new HopException( "Following required files are not accessible: " + message );
    }
  }

  public boolean init(){

    if ( super.init() ) {

      try {
        // Create the output row meta-data
        data.outputRowMeta = new RowMeta();
        meta.getFields( data.outputRowMeta, getTransformName(), null, null, this, metadataProvider ); // get the
        // metadata
        // populated
        data.nrTransformFields = data.outputRowMeta.size();

        if ( !meta.isFileField() ) {
          data.files = meta.getFileList( this );
          data.filessize = data.files.nrOfFiles();
          handleMissingFiles();
        } else {
          data.filessize = 0;
        }

      } catch ( Exception e ) {
        logError( "Error initializing transform: " + e.toString() );
        logError( Const.getStackTracker( e ) );
        return false;
      }

      data.filenr = 0;
      data.totalpreviousfields = 0;

      try {
				sampleRows = Integer.parseInt(environmentSubstitute(String.valueOf(meta.getSampleRows())));
				logDebug("Received StartRow: " + startRow + " SampleRows: " + sampleRows);
			} catch (NumberFormatException nfe) {
				logError("SampleRows couldn't be parsed");
				logDebug("SampleRow: " + meta.getSampleRows());
				return false;
			}

      return true;

    }
    return false;
  }

  public void dispose(){

    if ( data.file != null ) {
      try {
        data.file.close();
        data.file = null;
      } catch ( Exception e ) {
        // Ignore close errors
      }

    }
    super.dispose();
  }

}
