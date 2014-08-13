package br.com.giorgetti.util.extrato.service;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import br.com.giorgetti.util.extrato.dto.Expense;

public class ItauXlsReader extends XlsReader {

	private static final int IGNORE_ROW_COUNT = 11;
	private static final int DATE_COL = 2; // B
	private static final int DESC_COL = 5; // E
	private static final int VALU_COL = 6; // F

	@Override
	public List<Expense> generateExpenseList() {
		
		HSSFSheet sheet = wb.getSheetAt(wb.getActiveSheetIndex());
		Iterator<Row> rowIt = sheet.rowIterator();
		List<Expense> res = new ArrayList<Expense>();

		int index=0;
		while ( rowIt.hasNext() ) {
			Row row = rowIt.next();
			if ( ++index < IGNORE_ROW_COUNT )
				continue;
			
			Expense e = new Expense();
			Cell date  = row.getCell(DATE_COL);
			Cell desc  = row.getCell(DESC_COL);
			Cell value = row.getCell(VALU_COL);
			
			e.setDateTime(date.getStringCellValue());
			e.setDescription(desc.getStringCellValue());
			e.setValue(value.getNumericCellValue());
			
			res.add(e);
			
		}
		
		return res;
		
	}

}
