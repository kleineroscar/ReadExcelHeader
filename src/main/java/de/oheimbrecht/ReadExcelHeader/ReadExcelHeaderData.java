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
import org.apache.hop.core.fileinput.FileInputList;
import org.apache.hop.core.playlist.IFilePlayList;
import org.apache.hop.core.row.IRowMeta;
import org.apache.hop.pipeline.transform.BaseTransformData;
import org.apache.hop.pipeline.transform.ITransformData;
import org.apache.hop.pipeline.transform.errorhandling.IFileErrorHandler;

public class ReadExcelHeaderData extends BaseTransformData implements ITransformData {
	public RowMetaInterface outputRowMeta;
	public Object[] readrow;
	public int filenr;
	public FileInputList files;
	public FileObject file;
	public RowMetaInterface inputRowMeta;
	public int totalpreviousfields;
	public int indexOfFilenameField;
	public int indexOfWildcardField;
	public int indexOfExcludeWildcardField;
	public int filessize;
	public int nrTransformFields;
	public int nrLinesOnPage;
	public int nr_repeats;
	public Object[] previous_row;

	public ReadExcelHeaderData() {
			super();

			readrow = null;
			totalpreviousfields = 0;
			indexOfFilenameField = -1;
			indexOfWildcardField = -1;
			indexOfExcludeWildcardField = -1;
			
			nrTransformFields = 0;
			nrLinesOnPage = 0;

			filessize = 0;
			filenr = 0;

			nr_repeats = 0;
    		previous_row = null;
		}
}
