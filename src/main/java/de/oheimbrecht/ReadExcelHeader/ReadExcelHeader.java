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
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import org.pentaho.di.core.RowSet;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleValueException;
import org.pentaho.di.core.row.RowDataUtil;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.core.row.value.ValueMetaString;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStep;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;
import org.xml.sax.SAXException;

public class ReadExcelHeader extends BaseStep implements StepInterface {
	
	private String environmentFilename;
	private String realFilename;
	private int fieldnr;
	private int startRow;
	private ReadExcelHeaderData data;
	private ReadExcelHeaderMeta meta;
	
	
	public ReadExcelHeader(StepMeta stepMeta, StepDataInterface stepDataInterface, int copyNr, TransMeta transMeta,
			Trans trans) {
		super(stepMeta, stepDataInterface, copyNr, transMeta, trans);
	}
	
	
	public boolean init(StepMetaInterface smi, StepDataInterface sdi) {
		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;
		return super.init(smi, sdi);
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
			// clone the input row structure and place it in our data object
			data.previousRowMeta = getInputRowMeta().clone();
			data.outputRowMeta = (RowMetaInterface) getInputRowMeta().clone();
//			for (int i=0;i<data.outputRowMeta.getValueMetaList().size();i++) {
//				data.outputRowMeta.removeValueMeta(i);
//			}
			data.outputRowMeta.clear();
			data.outputRowMeta.addValueMeta(0, new ValueMetaString("workbookName"));
			data.outputRowMeta.addValueMeta(1, new ValueMetaString("sheetName"));
			data.outputRowMeta.addValueMeta(2, new ValueMetaString("columnName"));
			data.outputRowMeta.addValueMeta(3, new ValueMetaString("columnType"));
			data.outputRowMeta.addValueMeta(3, new ValueMetaString("columnDataFormat"));

			// use meta.getFields() to change it, so it reflects the output row structure
			meta.getFields(data.outputRowMeta, getStepname(), null, null, this, null, null);
//			data.fieldnr = data.previousRowMeta.indexOfValue("filename");//meta.getFilenameField());
			fieldnr = Integer.parseInt(meta.getFilenameField());
			startRow = Integer.parseInt(meta.getStartRow());
			if (fieldnr < 0) {
				throw new KettleValueException("CouldNotFindField :" + data.previousRowMeta.getFieldNames()[Integer.parseInt(meta.getFilenameField())].toString());
			}
		}
		environmentFilename = data.previousRowMeta.getString(r, fieldnr);
		try {
			URL url = new URL(environmentFilename);
			realFilename = url.getPath();
		} catch (MalformedURLException murl) {
			realFilename = environmentFilename;
		}
		// generate output row, make it correct size
		Object[] outputRow = RowDataUtil.resizeArray(r, data.outputRowMeta.size());

		InputStream file1InputStream = null;
		XSSFWorkbook workbook1 = null;
		try {
			file1InputStream = new FileInputStream(new File(realFilename));
			workbook1 = new XSSFWorkbook(file1InputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			XSSFSheet sheet = workbook1.getSheetAt(i);
			XSSFRow row = sheet.getRow(startRow);
			for (Cell myCell : row) {
				try {
				outputRow[0] = realFilename.toString();
				outputRow[1] = workbook1.getSheetName(i);
				outputRow[2] = myCell.toString();
				outputRow[3] = sheet.getRow(startRow + 1).getCell(myCell.getColumnIndex()).getCellTypeEnum();
				outputRow[4] = sheet.getRow(startRow + 1).getCell(myCell.getColumnIndex()).getCellStyle().getDataFormatString();
				} catch (Exception e) {
					log.logBasic(e.getMessage());
				}
				// put the row to the output row stream
				putRow(data.outputRowMeta, outputRow);
			}

			
			

			// log progress if it is time to to so
			if (checkFeedback(getLinesRead())) {
				logBasic("Processed Rows: " + getLinesRead()); // Some basic logging
			}
		}
		// indicate that processRow() should be called again
		return true;
	}

	public void dispose(StepMetaInterface smi, StepDataInterface sdi) {

		// Casting to step-specific implementation classes is safe
		meta = (ReadExcelHeaderMeta) smi;
		data = (ReadExcelHeaderData) sdi;

		// Add any step-specific initialization that may be needed here

		// Call superclass dispose()
		super.dispose(smi, sdi);
	}

	private void addFieldstoRowMeta(RowMetaInterface r, String origin) {
		ValueMetaInterface vWorkbook = new ValueMetaString("workbookName");

		// setting trim type to "both"
		vWorkbook.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// the name of the step that adds this field
		vWorkbook.setOrigin(origin);

		// modify the row structure and add the field this step generates
		r.addValueMeta(vWorkbook);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vSheet = new ValueMetaString("sheetName");

		// setting trim type to "both"
		vSheet.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// the name of the step that adds this field
		vSheet.setOrigin(origin);

		// modify the row structure and add the field this step generates
		r.addValueMeta(vSheet);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vColumnName = new ValueMetaString("columnName");

		// setting trim type to "both"
		vColumnName.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// the name of the step that adds this field
		vColumnName.setOrigin(origin);

		// modify the row structure and add the field this step generates
		r.addValueMeta(vColumnName);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vColumnType = new ValueMetaString("columnType");

		// setting trim type to "both"
		vColumnType.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// the name of the step that adds this field
		vColumnType.setOrigin(origin);

		// modify the row structure and add the field this step generates
		r.addValueMeta(vColumnType);
		
		// a value meta object contains the meta data for a field
		ValueMetaInterface vDataFormat = new ValueMetaString("columnDataFormat");

		// setting trim type to "both"
		vDataFormat.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// the name of the step that adds this field
		vDataFormat.setOrigin(origin);

		// modify the row structure and add the field this step generates
		r.addValueMeta(vDataFormat);
	}
}
