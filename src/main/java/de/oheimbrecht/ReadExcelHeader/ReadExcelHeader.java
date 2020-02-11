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
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.*;

import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleValueException;
import org.pentaho.di.core.row.RowDataUtil;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStep;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;

public class ReadExcelHeader extends BaseStep implements StepInterface {
	
	private String environmentFilename;
	private String realFilename;
	private int fieldnr;
	private int startRow;
	private ReadExcelHeaderMeta meta;
	private ReadExcelHeaderData data;
	
	
	InputStream file1InputStream = null;
	XSSFWorkbook workbook1 = null;
	
	
	public ReadExcelHeader(StepMeta stepMeta, StepDataInterface stepDataInterface, int copyNr, TransMeta transMeta,
			Trans trans) {
		super(stepMeta, stepDataInterface, copyNr, transMeta, trans);
	}
	
	
	public boolean init(StepMetaInterface smi, StepDataInterface sdi) {
		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;
		if (super.init(smi, sdi)) {
			first = true;
			return true;
		} 
		return false;
	}

	public boolean processRow(StepMetaInterface smi, StepDataInterface sdi) throws KettleException {

		// safely cast the step settings (meta) and runtime info (data) to specific
		// implementations
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		// get incoming row, getRow() potentially blocks waiting for more rows, returns
		// null if no more rows expected
		Object[] r = getRow();

		// if no more rows are expected, indicate step is finished and processRow()
		// should not be called again
		if (r == null) {
			setOutputDone();
			return false;
		}

		// the "first" flag is inherited from the base step implementation
		// it is used to guard some processing tasks, like figuring out field indexes
		// in the row structure that only need to be done once
		if (first) {
			first = false;
			data.outputRowMeta = (RowMetaInterface) getInputRowMeta().clone();

			// use meta.getFields() to change it, so it reflects the output row structure
			meta.getFields(data.outputRowMeta, getStepname(), null, null, this, null, null);
			try {
				fieldnr = Integer.parseInt(meta.getFilenameField());
				startRow = Integer.parseInt(meta.getStartRow());
			} catch (Exception e) {
				throw new KettleValueException("An error occurred while parsing the step settings.");
			}
			if (fieldnr < 0) {
				throw new KettleValueException("CouldNotFindField :" + getInputRowMeta().getFieldNames()[Integer.parseInt(meta.getFilenameField())].toString());
			}
		}
		environmentFilename = getInputRowMeta().getString(r, fieldnr);
		try {
			URL url = new URL(environmentFilename);
			realFilename = url.getPath();
		} catch (MalformedURLException murl) {
			realFilename = environmentFilename;
		}
		// generate output row, make it correct size
		Object[] outputRow = RowDataUtil.allocateRowData(5);

		try {
			file1InputStream = new FileInputStream(new File(realFilename));
			workbook1 = new XSSFWorkbook(file1InputStream);
		} catch (IOException e) {
			throw new KettleValueException("Could not parse the workbook");
		}
		
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			XSSFSheet sheet;
			XSSFRow row;
			
			try {
				sheet = workbook1.getSheetAt(i);
				row = sheet.getRow(startRow);
				log.logDebug("Found a sheet with the corresponding header row (from/to): " + row.getFirstCellNum() + "/" + row.getLastCellNum());
			} catch (Exception e) {
				log.logError("Unable to read sheet or row.\n Maybe the row given is empty.\n" + e.getMessage());
				throw new KettleValueException("Could not read sheet.");
			}
			for (short j=row.getFirstCellNum();j<row.getLastCellNum();j++) {
				try {
					log.logDebug("Processing the next cell with number: " + j);
					XSSFCell cellBelow = sheet.getRow(startRow + 1).getCell(row.getCell(j).getColumnIndex());
					log.logDebug("Created cellBelow");
					outputRow[0] = realFilename.toString();
					log.logRowlevel("Got workbook name: " + outputRow[0]);
					outputRow[1] = workbook1.getSheetName(i);
					log.logRowlevel("Got sheet name: " + outputRow[1]);
					outputRow[2] = row.getCell(j).toString();
					log.logRowlevel("Got cell header: " + outputRow[2]);
					try {
						outputRow[3] = cellBelow.getCellTypeEnum();
						log.logRowlevel("Got cell type from cellBelow: " + outputRow[3]);
						outputRow[4] = cellBelow.getCellStyle().getDataFormatString();
						log.logRowlevel("Got cell data format from cellBelow: " + outputRow[4]);
					} catch (Exception e) {
						outputRow[3] = "No data";
						outputRow[4] = "No data";
					}
				} catch (Exception e) {
					log.logError("Some error while getting the values. With i=" + String.valueOf(i) + "and j=" + String.valueOf(j));
					
					log.logError(e.getMessage());
					throw new KettleException(e.getMessage());
				}
				// put the row to the output row stream
				log.logDebug("Created the following row: " + Arrays.toString(outputRow));
				putRow(data.outputRowMeta, outputRow);
			}
			// log progress if it is time to to so
			if (checkFeedback(getLinesRead())) {
				logBasic("Processed Rows: " + getLinesRead()); // Some basic logging
			}
		}
		
		try {
			workbook1.close();
			file1InputStream.close();
		} catch (Exception e) {
			log.logError("Could not dispose workbook.\n" + e.getMessage());
		}

		// indicate that processRow() should be called again
		return true;
	}

	public void dispose(StepMetaInterface smi, StepDataInterface sdi) {

		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		// Add any step-specific initialization that may be needed here
//		fieldnr = -1;
//		startRow = -1;
//		realFilename = null;
//		environmentFilename = null;


		// Call superclass dispose()
		super.dispose(smi, sdi);
	}
}
