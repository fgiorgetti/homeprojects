package br.com.giorgetti.util.extrato.service;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import br.com.giorgetti.util.extrato.dto.Expense;
import br.com.giorgetti.util.extrato.exceptions.ExtratoParserException;
import br.com.giorgetti.util.extrato.exceptions.InvalidFileException;
import br.com.giorgetti.util.extrato.exceptions.ProcessingException;

public abstract class XlsReader extends ArrayList<ArrayList<String>>{

	private static final long serialVersionUID = -146132524350585527L;
	private Logger log = LoggerFactory.getLogger(getClass());
	protected HSSFWorkbook wb;

	public void loadFromXlsFile(String xlsFileName) throws ExtratoParserException {
		loadFromXlsFile(new File(xlsFileName));
	}
	
	public void loadFromXlsFile(File xlsFile) throws ExtratoParserException {
		
		if ( xlsFile == null || !xlsFile.exists() || 
				!xlsFile.canRead() ) {
			log.error("Unable to read log file");
		}
		
		BufferedInputStream is = null;
		try {
			is = new BufferedInputStream(new FileInputStream(xlsFile));
			wb = new HSSFWorkbook(is);
		} catch (FileNotFoundException e) {
			
			log.error("Could not find XLS file", e);
			throw new InvalidFileException();
		} catch (IOException e) {
			log.error("Error processing XLS file", e);
			throw new ProcessingException();
		}
		
	}
	
	public abstract List<Expense> generateExpenseList();
	
}
