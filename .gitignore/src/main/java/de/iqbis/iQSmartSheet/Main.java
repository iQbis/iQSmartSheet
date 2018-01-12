/*
// Copyright (c) 2018 M. Koetter, iQbis consulting GmbH
// Lahnstr. 35, 45478 Mülheim an der Ruhr, Germany
// www.iqbis.de
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
*/

package de.iqbis.iQSmartSheet;

import java.lang.*;

import java.io.File;
import java.io.FileReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Map;
import java.util.List;
import java.util.HashMap;
import java.util.Date;
import java.util.ArrayList;
import java.util.Iterator;

import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import net.sf.jxls.transformer.XLSTransformer;

public class Main {
	private static DateFormat dateFormat;

	public static void main(String args[]) {
		double nanos = (double) System.nanoTime();

		try {
			dateFormat = new SimpleDateFormat("yyyy-MM-dd");
			
			if (args.length != 3) {
				System.out.println("iQSmartSheet V1.0, 2018-01-12");
				System.out.println("Required arguments: <JSON_INPUT> <EXCEL_INPUT> <EXCEL_OUTPUT>");
				System.exit(1);
			}
			
			String json_inp = args[0];
			String excl_tpl = args[1];
			String excl_out = args[2];
	
			HashMap<String, Object> context = new HashMap<String, Object>();
			JSONParser parser = new JSONParser();
			JSONArray queries = (JSONArray) parser.parse(new FileReader(json_inp));

			for (Object obj : queries) {
				JSONObject query = (JSONObject) obj;
				String queryName = (String) query.get("query");
				JSONArray columns = (JSONArray) query.get("columns");
				JSONArray types = (JSONArray) query.get("types");
				JSONArray rows = (JSONArray) query.get("rows");
				ArrayList<HashMap> converted = new ArrayList<HashMap>();

				for (Object row : rows) {
					converted.add(convertRow(columns, types, (JSONArray) row));
				}

                context.put(queryName, converted);
			}

			XLSTransformer transformer = new XLSTransformer();
            transformer.transformXLS(excl_tpl, context, excl_out);
		} catch (Exception e) {
			e.printStackTrace();
			System.exit(2);
		}
		
		nanos = System.nanoTime() - nanos;
		System.out.printf("Completed succesfully, Time elapsed: %.3f sec", nanos / 1e9);
	}

	private static HashMap convertRow(JSONArray columns, JSONArray types, JSONArray row) throws java.text.ParseException {
		HashMap<String, Object> rowObject = new HashMap<String, Object>();

		for (int i = 0; i < columns.size(); i++) {
			rowObject.put((String) columns.get(i), convertValue((String) row.get(i), (String) types.get(i)));
		}

		return rowObject;
	}

	private static Object convertValue(String value, String type) throws java.text.ParseException {
		if (value == null) {
			return "";
		} else if (type.equals("DATE")) {
			return dateFormat.parse(value);
		} else if (type.equals("LONG")) {
			return new Long(value);
		} else if (type.equals("NUMBER")) {
			return new Double(value);
		}

		return value;
	}
}
