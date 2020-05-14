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

import java.util.List;
import java.util.Map;

import org.eclipse.swt.widgets.Shell;
import org.pentaho.di.core.CheckResult;
import org.pentaho.di.core.CheckResultInterface;
import org.pentaho.di.core.Const;
import org.pentaho.di.core.Counter;
import org.pentaho.di.core.annotations.Step;
import org.pentaho.di.core.database.DatabaseMeta;
import org.pentaho.di.core.exception.KettleDatabaseException;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleStepException;
import org.pentaho.di.core.exception.KettleXMLException;
import org.pentaho.di.core.fileinput.FileInputList;
import org.pentaho.di.core.injection.Injection;
import org.pentaho.di.core.injection.InjectionSupported;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.core.row.value.ValueMetaString;
import org.pentaho.di.core.variables.VariableSpace;
import org.pentaho.di.core.xml.XMLHandler;
import org.pentaho.di.repository.ObjectId;
import org.pentaho.di.repository.Repository;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStepMeta;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepDialogInterface;
import org.pentaho.di.trans.step.StepInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;
import org.pentaho.metastore.api.IMetaStore;
import org.w3c.dom.Node;

@Step(id = "ReadExcelHeader", image = "REH.svg", i18nPackageName = "de.oheimbrecht.ReadExcelHeader", name = "ReadExcelHeader.Step.Name", description = "ReadExcelHeader.Step.Description", categoryDescription = "i18n:org.pentaho.di.trans.step:BaseStep.Category.Utility", documentationUrl = "https://github.com/kleineroscar/ReadExcelHeader")

@InjectionSupported(localizationPrefix = "ReadExcelHeader.Injection.")
public class ReadExcelHeaderMeta extends BaseStepMeta implements StepMetaInterface {
	/** Field to read from */
	@Injection(name = "FIELD_OF_FILENAMES")
	private String filenameField;
	@Injection(name = "ROW_TO_START_ON")
	private String startRow;
	@Injection(name = "NUMBER_OF_ROWS_TO_SAMPLE")
	private String sampleRows;

	private String startRowField;

	private boolean filefield;
	private boolean startrowfield;

	public static final String[] RequiredFilesDesc = new String[] { Messages.getString("System.Combo.No"),
			Messages.getString("System.Combo.Yes") };
	public static final String[] RequiredFilesCode = new String[] { "N", "Y" };
	private static final String NO = "N";
	private static final String YES = "Y";

	/** Array of filenames */
	private String[] fileName;

	/** Wildcard or filemask (regular expression) */
	private String[] fileMask;

	/** Wildcard or filemask to exclude (regular expression) */
	private String[] excludeFileMask;

	/** Flag indicating that a row number field should be included in the output */
	private boolean includeFilesCount;

	/** Array of boolean values as string, indicating if a file is required. */
	private String[] fileRequired;

	/**
	 * Array of boolean values as string, indicating if we need to fetch sub
	 * folders.
	 */
	private String[] includeSubFolders;

	public ReadExcelHeaderMeta() {
		super();
	}

	/**
	 * Called by Spoon to get a new instance of the SWT dialog for the step. A
	 * standard implementation passing the arguments to the constructor of the step
	 * dialog is recommended.
	 * 
	 * @param shell     an SWT Shell
	 * @param meta      description of the step
	 * @param transMeta description of the the transformation
	 * @param name      the name of the step
	 * @return new instance of a dialog for this step
	 */
	public StepDialogInterface getDialog(Shell shell, StepMetaInterface meta, TransMeta transMeta, String name) {
		return new ReadExcelHeaderDialog(shell, meta, transMeta, name);
	}

	/**
	 * Called by PDI to get a new instance of the step implementation. A standard
	 * implementation passing the arguments to the constructor of the step class is
	 * recommended.
	 * 
	 * @param stepMeta          description of the step
	 * @param stepDataInterface instance of a step data class
	 * @param cnr               copy number
	 * @param transMeta         description of the transformation
	 * @param disp              runtime implementation of the transformation
	 * @return the new instance of a step implementation
	 */
	public StepInterface getStep(StepMeta stepMeta, StepDataInterface stepDataInterface, int cnr, TransMeta transMeta,
			Trans disp) {
		return new ReadExcelHeader(stepMeta, stepDataInterface, cnr, transMeta, disp);
	}

	/**
	 * Called by PDI to get a new instance of the step data class.
	 */
	public StepDataInterface getStepData() {
		return new ReadExcelHeaderData();
	}

	/**
	 * This method is called every time a new step is created and should
	 * allocate/set the step configuration to sensible defaults. The values set here
	 * will be used by Spoon when a new step is created.
	 */
	public void setDefault() {
		filenameField = "";
		startRow = "0";
		sampleRows = "10";
		filefield = false;
		startrowfield = false;

		int nrFiles = 0;

		allocate(nrFiles);

		for (int i = 0; i < nrFiles; i++) {
			fileName[i] = "filename" + (i + 1);
			fileMask[i] = "";
			excludeFileMask[i] = "";
			fileRequired[i] = RequiredFilesCode[0];
			includeSubFolders[i] = RequiredFilesCode[0];
		}
	}

	// @Override
	// public boolean excludeFromCopyDistributeVerification()
	// {
	// return true;
	// }

	/**
	 * This method is used when a step is duplicated in Spoon. It needs to return a
	 * deep copy of this step meta object. Be sure to create proper deep copies if
	 * the step configuration is stored in modifiable objects.
	 * 
	 * See org.pentaho.di.trans.steps.rowgenerator.RowGeneratorMeta.clone() for an
	 * example on creating a deep copy.
	 * 
	 * @return a deep copy of this
	 */
	public Object clone() {
		ReadExcelHeaderMeta retval = (ReadExcelHeaderMeta) super.clone();
		retval.filenameField = filenameField;
		retval.startRow = startRow;
		retval.sampleRows = sampleRows;

		int nrFiles = fileName.length;

		retval.allocate(nrFiles);
		System.arraycopy(fileName, 0, retval.fileName, 0, nrFiles);
		System.arraycopy(fileMask, 0, retval.fileMask, 0, nrFiles);
		System.arraycopy(excludeFileMask, 0, retval.excludeFileMask, 0, nrFiles);
		System.arraycopy(fileRequired, 0, retval.fileRequired, 0, nrFiles);
		System.arraycopy(includeSubFolders, 0, retval.includeSubFolders, 0, nrFiles);

		return retval;
	}

	// For compatibility with 7.x
	@Override
	public String getDialogClassName() {
		return ReadExcelHeaderDialog.class.getName();
	}

	/**
	 * This method is called to determine the changes the step is making to the
	 * row-stream. To that end a RowMetaInterface object is passed in, containing
	 * the row-stream structure as it is when entering the step. This method must
	 * apply any changes the step makes to the row stream. Usually a step adds
	 * fields to the row-stream.
	 * 
	 * @param inputRowMeta the row structure coming in to the step
	 * @param name         the name of the step making the changes
	 * @param info         row structures of any info steps coming in
	 * @param nextStep     the description of a step this step is passing rows to
	 * @param space        the variable space for resolving variables
	 * @param repository   the repository instance optionally read from
	 * @param metaStore    the metaStore to optionally read from
	 */
	public void getFields(RowMetaInterface inputRowMeta, String name, RowMetaInterface[] info, StepMeta nextStep,
			VariableSpace space, Repository repository, IMetaStore metaStore) throws KettleStepException {

		/*
		 * This implementation appends the outputField to the row-stream
		 */
		// a value meta object contains the meta data for a field
		ValueMetaInterface vWorkbook = new ValueMetaString("workbookName", 500, -1);

		// the name of the step that adds this field
		vWorkbook.setOrigin(name);

		vWorkbook.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// modify the row structure and add the field this step generates
		inputRowMeta.addValueMeta(vWorkbook);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vSheet = new ValueMetaString("sheetName", 500, -1);

		// the name of the step that adds this field
		vSheet.setOrigin(name);

		vSheet.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// modify the row structure and add the field this step generates
		inputRowMeta.addValueMeta(vSheet);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vColumnName = new ValueMetaString("columnName", 500, -1);

		// the name of the step that adds this field
		vColumnName.setOrigin(name);

		vColumnName.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// modify the row structure and add the field this step generates
		inputRowMeta.addValueMeta(vColumnName);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vColumnType = new ValueMetaString("columnType", 500, -1);

		// the name of the step that adds this field
		vColumnType.setOrigin(name);

		vColumnType.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// modify the row structure and add the field this step generates
		inputRowMeta.addValueMeta(vColumnType);

		// a value meta object contains the meta data for a field
		ValueMetaInterface vColumnDataFormat = new ValueMetaString("columnDataFormat", 500, -1);

		// the name of the step that adds this field
		vColumnDataFormat.setOrigin(name);

		vColumnDataFormat.setTrimType(ValueMetaInterface.TRIM_TYPE_BOTH);

		// modify the row structure and add the field this step generates
		inputRowMeta.addValueMeta(vColumnDataFormat);
	}

	public FileInputList getFiles(VariableSpace space) {
		return FileInputList.createFileList(space, fileName, fileMask, excludeFileMask, fileRequired,
				includeSubFolderBoolean());
	}

	private boolean[] includeSubFolderBoolean() {
		int len = fileName.length;
		boolean[] includeSubFolderBoolean = new boolean[len];
		for (int i = 0; i < len; i++) {
			includeSubFolderBoolean[i] = YES.equalsIgnoreCase(includeSubFolders[i]);
		}
		return includeSubFolderBoolean;
	}

	/**
	 * This method is called when the user selects the "Verify Transformation"
	 * option in Spoon. A list of remarks is passed in that this method should add
	 * to. Each remark is a comment, warning, error, or ok. The method should
	 * perform as many checks as necessary to catch design-time errors.
	 * 
	 * Typical checks include: - verify that all mandatory configuration is given -
	 * verify that the step receives any input, unless it's a row generating step -
	 * verify that the step does not receive any input if it does not take them into
	 * account - verify that the step finds fields it relies on in the row-stream
	 * 
	 * @param remarks   the list of remarks to append to
	 * @param transMeta the description of the transformation
	 * @param stepMeta  the description of the step
	 * @param prev      the structure of the incoming row-stream
	 * @param input     names of steps sending input to the step
	 * @param output    names of steps this step is sending output to
	 * @param info      fields coming in from info steps
	 * @param metaStore metaStore to optionally read from
	 */
	public void check(List<CheckResultInterface> remarks, TransMeta transMeta, StepMeta stepMeta, RowMetaInterface prev,
			String[] input, String[] output, RowMetaInterface info, VariableSpace space, Repository repository,
			IMetaStore metaStore) {
		CheckResult cr;

		if (prev == null || prev.size() == 0) {
			cr = new CheckResult(CheckResult.TYPE_RESULT_WARNING, "Not receiving any fields from previous steps!",
					stepMeta);
			remarks.add(cr);
		} else {
			cr = new CheckResult(CheckResult.TYPE_RESULT_OK,
					"Step is connected to previous one, receiving " + prev.size() + " fields", stepMeta);
			remarks.add(cr);
		}

		// See if we have input streams leading to this step!
		if (input.length > 0) {
			cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step is receiving info from other steps.", stepMeta);
			remarks.add(cr);
		} else {
			cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No input received from other steps!", stepMeta);
			remarks.add(cr);
		}

		if (!filenameField.isEmpty()) {
			cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received a filename.", stepMeta);
			remarks.add(cr);
		} else {
			cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No filename received!", stepMeta);
			remarks.add(cr);
		}

		if (!startRow.isEmpty()) {
			cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received a row to start on.", stepMeta);
			remarks.add(cr);
		} else {
			cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No start row received!", stepMeta);
			remarks.add(cr);
		}

		if (!sampleRows.isEmpty()) {
			cr = new CheckResult(CheckResult.TYPE_RESULT_OK, "Step received number of rows to sample on.", stepMeta);
			remarks.add(cr);
		} else {
			cr = new CheckResult(CheckResult.TYPE_RESULT_ERROR, "No number of sample rows received!", stepMeta);
			remarks.add(cr);
		}
	}

	public String getFilenameField() {
		return filenameField;
	}

	public void setFilenameField(final String filenameField) {
		this.filenameField = filenameField;
	}

	public String getStartRowField() {
		return startRowField;
	}

	public void setStartRowField(final String startRowField) {
		this.startRowField = startRowField;
	}

	public String getStartRow() {
		return startRow;
	}

	public void setStartRow(final String StartRow) {
		this.startRow = StartRow;
	}

	public String getSampleRows() {
		return sampleRows;
	}

	public void setSampleRows(String sampleRows) {
		this.sampleRows = sampleRows;
	}

	public String getXML() {
		StringBuilder retval = new StringBuilder(500);
		retval.append("   " + XMLHandler.addTagValue("filenamefield", filenameField));
		retval.append("   " + XMLHandler.addTagValue("startrow", startRow));
		retval.append("   " + XMLHandler.addTagValue("sampleRows", sampleRows));
		retval.append("    ").append(XMLHandler.addTagValue("filefield", filefield));
		retval.append("    ").append(XMLHandler.addTagValue("startrowfield", startrowfield));

		retval.append("    <file>").append(Const.CR);
		for (int i = 0; i < fileName.length; i++) {
			retval.append("      ").append(XMLHandler.addTagValue("name", fileName[i]));
			retval.append("      ").append(XMLHandler.addTagValue("filemask", fileMask[i]));
			retval.append("      ").append(XMLHandler.addTagValue("exclude_filemask", excludeFileMask[i]));
			retval.append("      ").append(XMLHandler.addTagValue("file_required", fileRequired[i]));
			retval.append("      ").append(XMLHandler.addTagValue("include_subfolders", includeSubFolders[i]));
			parentStepMeta.getParentTransMeta().getNamedClusterEmbedManager().registerUrl(fileName[i]);
		}
		retval.append("    </file>").append(Const.CR);

		return retval.toString();
	}

	public void loadXML(Node stepnode, List<DatabaseMeta> databases, Map<String, Counter> counters)
			throws KettleXMLException {
		try {
			filenameField = XMLHandler.getTagValue(stepnode, "filenamefield");
			startRow = XMLHandler.getTagValue(stepnode, "startrow");
			sampleRows = XMLHandler.getTagValue(stepnode, "sampleRows");

			filefield = "Y".equalsIgnoreCase(XMLHandler.getTagValue(stepnode, "filefield"));
			startrowfield = "Y".equalsIgnoreCase(XMLHandler.getTagValue(stepnode, "startrowfield"));

			Node filenode = XMLHandler.getSubNode(stepnode, "file");
			int nrFiles = XMLHandler.countNodes(filenode, "name");
			allocate(nrFiles);

			for (int i = 0; i < nrFiles; i++) {
				Node filenamenode = XMLHandler.getSubNodeByNr(filenode, "name", i);
				Node filemasknode = XMLHandler.getSubNodeByNr(filenode, "filemask", i);
				Node excludefilemasknode = XMLHandler.getSubNodeByNr(filenode, "exclude_filemask", i);
				Node fileRequirednode = XMLHandler.getSubNodeByNr(filenode, "file_required", i);
				Node includeSubFoldersnode = XMLHandler.getSubNodeByNr(filenode, "include_subfolders", i);
				fileName[i] = XMLHandler.getNodeValue(filenamenode);
				fileMask[i] = XMLHandler.getNodeValue(filemasknode);
				excludeFileMask[i] = XMLHandler.getNodeValue(excludefilemasknode);
				fileRequired[i] = XMLHandler.getNodeValue(fileRequirednode);
				includeSubFolders[i] = XMLHandler.getNodeValue(includeSubFoldersnode);
			}
		} catch (Exception e) {
			throw new KettleXMLException("Unable to read step info from XML node", e);
		}
	}

	public void readRep(Repository rep, ObjectId id_step, List<DatabaseMeta> databases, Map<String, Counter> counters)
			throws KettleException {
		try {
			filenameField = rep.getStepAttributeString(id_step, "filenamefield");
			startRow = rep.getStepAttributeString(id_step, "startrow");
			sampleRows = rep.getStepAttributeString(id_step, "sampleRows");

			filefield = rep.getStepAttributeBoolean(id_step, "filefield");
			startrowfield= rep.getStepAttributeBoolean(id_step, "startrowfield");

			int nrFiles = rep.countNrStepAttributes(id_step, "file_name");

			allocate(nrFiles);

			for (int i = 0; i < nrFiles; i++) {
				fileName[i] = rep.getStepAttributeString(id_step, i, "file_name");
				fileMask[i] = rep.getStepAttributeString(id_step, i, "file_mask");
				excludeFileMask[i] = rep.getStepAttributeString(id_step, i, "exclude_file_mask");
				fileRequired[i] = rep.getStepAttributeString(id_step, i, "file_required");
				if (!YES.equalsIgnoreCase(fileRequired[i])) {
					fileRequired[i] = NO;
				}
				includeSubFolders[i] = rep.getStepAttributeString(id_step, i, "include_subfolders");
				if (!YES.equalsIgnoreCase(includeSubFolders[i])) {
					includeSubFolders[i] = NO;
				}
			}
		} catch (KettleDatabaseException dbe) {
			throw new KettleException("error reading step with id_step=" + id_step + " from the repository", dbe);
		} catch (Exception e) {
			throw new KettleException("Unexpected error reading step with id_step=" + id_step + " from the repository",
					e);
		}
	}

	public void saveRep(Repository rep, ObjectId id_transformation, ObjectId id_step) throws KettleException {
		try {
			rep.saveStepAttribute(id_transformation, id_step, "filenamefield", filenameField);
			rep.saveStepAttribute(id_transformation, id_step, "startrow", startRow);
			rep.saveStepAttribute(id_transformation, id_step, "sampleRows", sampleRows);
		} catch (KettleDatabaseException dbe) {
			throw new KettleException("Unable to save step information to the repository, id_step=" + id_step, dbe);
		}
	}

	public void allocate(int nrfiles) {
		fileName = new String[nrfiles];
		fileMask = new String[nrfiles];
		excludeFileMask = new String[nrfiles];
		fileRequired = new String[nrfiles];
		includeSubFolders = new String[nrfiles];
	}

	/**
	 * @return Returns the excludeFileMask.
	 */
	public String[] getExcludeFileMask() {
		return excludeFileMask;
	}

	/**
	 * @param excludeFileMask The excludeFileMask to set.
	 */
	public void setExcludeFileMask(String[] excludeFileMask) {
		this.excludeFileMask = excludeFileMask;
	}

	/**
	 * @return Returns the output filename_Field.
	 */
	public String getFileNameField() {
		return filenameField;
	}

	/**
	 * @param filenameField The output filename_field to set.
	 */
	public void setFileNameField(String filenameField) {
		this.filenameField = filenameField;
	}

	/**
	 * @return Returns the fileMask.
	 */
	public String[] getFileMask() {
		return fileMask;
	}

	public void setFileRequired(String[] fileRequiredin) {
		for (int i = 0; i < fileRequiredin.length; i++) {
			this.fileRequired[i] = getRequiredFilesCode(fileRequiredin[i]);
		}
	}

	public String[] getIncludeSubFolders() {
		return includeSubFolders;
	}

	public void setIncludeSubFolders(String[] includeSubFoldersin) {
		for (int i = 0; i < includeSubFoldersin.length; i++) {
			this.includeSubFolders[i] = getRequiredFilesCode(includeSubFoldersin[i]);
		}
	}

	public String getRequiredFilesCode(String tt) {
		if (tt == null) {
			return RequiredFilesCode[0];
		}
		if (tt.equals(RequiredFilesDesc[1])) {
			return RequiredFilesCode[1];
		} else {
			return RequiredFilesCode[0];
		}
	}

	public String getRequiredFilesDesc(String tt) {
		if (tt == null) {
			return RequiredFilesDesc[0];
		}
		if (tt.equals(RequiredFilesCode[1])) {
			return RequiredFilesDesc[1];
		} else {
			return RequiredFilesDesc[0];
		}
	}

	/**
	 * @param fileMask The fileMask to set.
	 */
	public void setFileMask(String[] fileMask) {
		this.fileMask = fileMask;
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
	public void setFileName(String[] fileName) {
		this.fileName = fileName;
	}

	/**
	 * @return Returns the includeCountFiles.
	 */
	public boolean includeCountFiles() {
		return includeFilesCount;
	}

	/**
	 * @param includeFilesCount The "includes files count" flag to set.
	 */
	public void setIncludeCountFiles(boolean includeFilesCount) {
		this.includeFilesCount = includeFilesCount;
	}

	public String[] getFileRequired() {
		return this.fileRequired;
	}

	/**
	 * @return Returns the File field.
	 */
	public boolean isFileField() {
		return filefield;
	}

	/**
	 * @param filefield The file field to set.
	 */
	public void setFileField(boolean filefield) {
		this.filefield = filefield;
	}

		/**
	 * @return Returns the start row field.
	 */
	public boolean isStartRowField() {
		return startrowfield;
	}

	/**
	 * @param startrowfield The start row field to set.
	 */
	public void setStartRowField(boolean startrowfield) {
		this.startrowfield = startrowfield;
	}
}
