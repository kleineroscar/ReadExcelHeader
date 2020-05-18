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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.*;
import org.pentaho.di.core.Const;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleStepException;
import org.pentaho.di.core.row.RowDataUtil;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.util.Utils;
import org.pentaho.di.core.vfs.KettleVFS;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStep;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;

public class ReadExcelHeader extends BaseStep {

	private int startRow;
	private int sampleRows;
	private ReadExcelHeaderMeta meta;
	private ReadExcelHeaderData data;

	InputStream file1InputStream = null;
	XSSFWorkbook workbook1 = null;

	private String filePath;

	public ReadExcelHeader(StepMeta stepMeta, StepDataInterface stepDataInterface, int copyNr, TransMeta transMeta,
			Trans trans) {
		super(stepMeta, stepDataInterface, copyNr, transMeta, trans);

		meta = (ReadExcelHeaderMeta) getStepMeta().getStepMetaInterface();
		data = (ReadExcelHeaderData) stepDataInterface;
	}

	public boolean init(StepMetaInterface smi, StepDataInterface sdi) {
		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		if (super.init(smi, sdi)) {
			first = true;

			// Set Embedded NamedCluter MetatStore Provider Key so that it can be passed to
			// VFS
			if (getTransMeta().getNamedClusterEmbedManager() != null) {
				getTransMeta().getNamedClusterEmbedManager().passEmbeddedMetastoreKey(this,
						getTransMeta().getEmbeddedMetastoreProviderKey());
			}

			if (!meta.isFileField()) {
				data.files = meta.getFiles(this);
				if (data.files == null || data.files.nrOfFiles() == 0) {
					logError(Messages.getString("ReadExcelHeaderDialog.Log.NoFiles"));
					return false;
				}
				try {
					// Create the output row meta-data
					data.outputRowMeta = new RowMeta();
					meta.getFields(data.outputRowMeta, getStepname(), null, null, this, repository, metaStore); // get
																												// the
																												// metadata
																												// populated

				} catch (Exception e) {
					logError("Error initializing step: " + e.toString());
					logError(Const.getStackTracker(e));
					return false;
				}
			}
			data.rownr = 0;
			data.filenr = 0;
			data.totalpreviousfields = 0;

			try {
				startRow = Integer.parseInt(environmentSubstitute(meta.getStartRow()));
				sampleRows = Integer.parseInt(environmentSubstitute(meta.getSampleRows()));
				logDebug("Received StartRow: " + startRow + " SampleRows: " + sampleRows);
			} catch (NumberFormatException nfe) {
				logError("StartRow or SampleRows couldn't be parsed");
				logDebug("Startrow: " + meta.getStartRow());
				logDebug("SampleRow: " + meta.getSampleRows());
				return false;
			}

			return true;
		}
		return false;
	}

	private Object[] getOneRow() throws KettleException {
		if (!openNextFile()) {
			return null;
		}

		// Build an empty row based on the meta-data
		Object[] r;
		try {
			// Create new row or clone
			if (meta.isFileField()) {
				r = data.readrow.clone();
				r = RowDataUtil.resizeArray(r, data.outputRowMeta.size());
			} else {
				r = RowDataUtil.allocateRowData(data.outputRowMeta.size());
			}

			r[data.totalpreviousfields] = data.rownr;

			incrementLinesInput();

		} catch (Exception e) {
			throw new KettleException("Unable to read row from file", e);
		}

		return r;
	}

	public boolean processRow(StepMetaInterface smi, StepDataInterface sdi) throws KettleException {

		// safely cast the step settings (meta) and runtime info (data) to specific
		// implementations
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		// get incoming row, getRow() potentially blocks waiting for more rows, returns
		// null if no more rows expected
		Object[] r;// = getRow();

		try {
			Object[] outputRowData = getOneRow();
			if (outputRowData == null) {
				setOutputDone(); // signal end to receiver(s)
				return false; // end of data or error.
			}
			r = outputRowData;

			// if ((!meta.isFileField() && data.last_file) || meta.isFileField()) {
			// putRow(data.outputRowMeta, outputRowData); // copy row to output rowset(s);
			// if (log.isDetailed()) {
			// logDetailed(BaseMessages.getString(PKG,
			// "GetFilesRowsCount.Log.TotalRowsFiles"), data.rownr,
			// data.filenr);
			// }
			// }
			
			logDebug("Is file field used? " + (meta.isFileField() ? "Yes" : "No"));
			logDebug("data.file is: " + data.file.toString());
			filePath = data.file.toString();
			
			getHeader(r);
			// indicate that processRow() should be called again
			return true;

		} catch (KettleException e) {

			logError(Messages.getString("ReadExcelHeaderDialog.ErrorInStepRunning", e.getMessage()));
			setErrors(1);
			stopAll();
			setOutputDone(); // signal end to receiver(s)
			return false;
		}
		// return true;

	}

	private boolean openNextFile() {
		if (data.last_file) {
			return false; // Done!
		}

		try {
			if (!meta.isFileField()) {
				if (data.filenr >= data.files.nrOfFiles()) {
					// finished processing!

					if (log.isDetailed()) {
						logDetailed(Messages.getString("ReadExcelHeaderDialog.Log.FinishedProcessing"));
					}
					return false;
				}

				// Is this the last file?
				data.last_file = (data.filenr == data.files.nrOfFiles() - 1);
				data.file = data.files.getFile((int) data.filenr);

			} else {
				data.readrow = getRow(); // Get row from input rowset & set row busy!
				if (data.readrow == null) {
					if (log.isDetailed()) {
						logDetailed(Messages.getString("ReadExcelHeaderDialog.Log.FinishedProcessing"));
					}
					return false;
				}

				if (first) {
					first = false;

					data.inputRowMeta = getInputRowMeta();
					data.outputRowMeta = data.inputRowMeta.clone();
					meta.getFields(data.outputRowMeta, getStepname(), null, null, this, repository, metaStore);

					// Get total previous fields
					data.totalpreviousfields = data.inputRowMeta.size();

					// Check is filename field is provided
					if (Utils.isEmpty(meta.getFileNameField())) {
						logError(Messages.getString("ReadExcelHeaderDialog.Log.NoField"));
						throw new KettleException(Messages.getString("ReadExcelHeaderDialog.Log.NoField"));
					}

					// cache the position of the field
					if (data.indexOfFilenameField < 0) {
						data.indexOfFilenameField = getInputRowMeta().indexOfValue(meta.getFileNameField());
						if (data.indexOfFilenameField < 0) {
							// The field is unreachable !
							logError(Messages.getString("ReadExcelHeaderDialog.Log.ErrorFindingField",
									meta.getFileNameField()));
							throw new KettleException(Messages.getString(
									"ReadExcelHeaderDialog.Exception.CouldnotFindField", meta.getFileNameField()));
						}
					}

				} // End if first

				String filename = getInputRowMeta().getString(data.readrow, data.indexOfFilenameField);
				if (log.isDetailed()) {
					logDetailed(Messages.getString("ReadExcelHeaderDialog.Log.FilenameInStream",
							meta.getFileNameField(), filename));
				}

				data.file = KettleVFS.getFileObject(filename, getTransMeta());

				// Init Row number
				if (meta.isFileField()) {
					data.rownr = 0;
				}
			}

			// Move file pointer ahead!
			data.filenr++;

			if (log.isDetailed()) {
				logDetailed(Messages.getString("ReadExcelHeaderDialog.Log.OpeningFile", data.file.toString()));
			}

			// if ( log.isDetailed() ) {
			// logDetailed( Messages.getString( "ReadExcelHeaderDialog.Log.FileOpened",
			// data.file.toString() ) );
			// }

		} catch (Exception e) {
			logError(Messages.getString("ReadExcelHeaderDialog.Log.UnableToOpenFile", "" + data.filenr,
					data.file.toString(), e.toString()));
			stopAll();
			setErrors(1);
			return false;
		}
		return true;
	}

	private void getHeader(Object[] r) throws KettleStepException {
		log.logDebug("Filepath is: " + filePath);
		filePath = (filePath.contains("file://") ? filePath.substring(7) : filePath);
		logDebug("cleansed filepath is: " + filePath);

		try {
			// file1InputStream = new URL(filePath).openStream();
			file1InputStream = new FileInputStream(new File(filePath));
		} catch (IOException e) {
			log.logDebug("Supplied file: " + filePath);
			log.logDebug(e.getMessage());
			throw new KettleStepException("Could not read the file provided.");
		}

		try {
			workbook1 = new XSSFWorkbook(file1InputStream);
		} catch (IOException e) {
			log.logDebug(e.getMessage());
			throw new KettleStepException("Could not parse the workbook from the file");
		}

		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			XSSFSheet sheet;
			XSSFRow row;

			try {
				sheet = workbook1.getSheetAt(i);
			} catch (Exception e) {
				log.logError("Unable to read sheet\n" + e.getMessage());
				throw new KettleStepException("Could not read sheet with number: " + i);
			}
			try {
				row = sheet.getRow(startRow);
				log.logDebug("Found a sheet with the corresponding header row (from/to): " + row.getFirstCellNum() + "/"
						+ row.getLastCellNum());
			} catch (Exception e) {
				log.logError("Unable to read row.\nMaybe the row given is empty.\n" + e.getMessage());
				throw new KettleStepException("Could not read row with startrow: " + startRow);
			}
			for (short j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				// generate output row, make it correct size
				Object[] outputRow = RowDataUtil.createResizedCopy(r, data.outputRowMeta.size());

				int lastMeta = data.outputRowMeta.size();
				try {
					log.logDebug("Processing the next cell with number: " + j);
					outputRow[lastMeta - 5] = (new File(filePath)).getName();
					log.logRowlevel("Got workbook name: " + outputRow[lastMeta - 5] + " setting in "
							+ String.valueOf(lastMeta - 5));
					outputRow[lastMeta - 4] = workbook1.getSheetName(i);
					log.logRowlevel("Got sheet name: " + outputRow[lastMeta - 4] + " setting in "
							+ String.valueOf(lastMeta - 4));
					outputRow[lastMeta - 3] = row.getCell(j).toString();
					log.logRowlevel("Got cell header: " + outputRow[lastMeta - 3] + " setting in "
							+ String.valueOf(lastMeta - 3));
				} catch (Exception e) {
					log.logError("Some error while getting the values. With Sheetnumber:" + String.valueOf(i)
							+ " and Column number:" + String.valueOf(j));

					log.logError(e.getMessage());
					throw new KettleStepException(e.getMessage());
				}

				Map<String, String[]> cellInfo = new HashMap<>();
				log.logDebug("Startrow is: " + startRow);
				log.logDebug("Samplerows is: " + sampleRows);
				for (int k = startRow + 1; k <= sampleRows; k++) {
					log.logDebug("Going into loop for getting cell info with k= " + k);
					try {
						XSSFCell cell = sheet.getRow(k).getCell(row.getCell(j).getColumnIndex());
						log.logDebug("Adding type and style to list: '" + cell.getCellTypeEnum().toString() + "/"
								+ cell.getCellStyle().getDataFormatString() + "'");
						cellInfo.put(cell.getCellTypeEnum().toString() + cell.getCellStyle().getDataFormatString(),
								new String[] { cell.getCellTypeEnum().toString(),
										cell.getCellStyle().getDataFormatString() });
					} catch (Exception e) {
						log.logDebug("Couldn't get Field info in row " + String.valueOf(k));
					}
				}
				if (cellInfo.size() == 0) {
					outputRow[lastMeta - 2] = "NO DATA";
					outputRow[lastMeta - 1] = "NO DATA";
				} else if (cellInfo.size() > 1) {
					outputRow[lastMeta - 2] = "STRING";
					outputRow[lastMeta - 1] = "Mixed";
				} else {
					String[] info = (String[]) cellInfo.values().toArray()[0];
					outputRow[lastMeta - 2] = info[0];
					outputRow[lastMeta - 1] = info[1];
				}

				// put the row to the output row stream
				log.logDebug(
						"Created the following row: " + Arrays.toString(outputRow) + " ;map size=" + cellInfo.size());
				putRow(data.outputRowMeta, outputRow);
			}
			// log progress if it is time to to so
			if (checkFeedback(getLinesRead())) {
				logBasic("Processed Rows: " + getLinesRead()); // Some basic logging
			}
		}
		try {
			workbook1.close();
			// file1InputStream.close();
		} catch (Exception e) {
			new KettleException("Could not dispose workbook.\n" + e.getMessage());
		}
		try {
			file1InputStream.close();
		} catch (IOException e) {
			new KettleException("Could not dispose FileInputStream.\n" + e.getMessage());
		}
	}

	public void dispose(StepMetaInterface smi, StepDataInterface sdi) {

		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		// // Add any step-specific initialization that may be needed here
		// try {
		// file1InputStream.close();
		// } catch (IOException e) {
		// log.logError("Could not dispose FileInputStream.\n" + e.getMessage());
		// }

		// Call superclass dispose()
		super.dispose(smi, sdi);
	}
}
