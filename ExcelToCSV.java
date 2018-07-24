package iwebz.utility.preprocessor;

import java.io.*;
import java.util.Iterator;
import java.util.HashMap;
import java.util.Map;
import java.util.Date;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import com.monitorjbl.xlsx.StreamingReader;

class ExcelToCSV {

  public HashMap convertXlsxToCSV(String infile, String outfile,String password) 
  {
    int rownum=0;
    int rowcount=0;
    HashMap msg_log=new HashMap();
    StringBuffer cellDData = new StringBuffer();
    
    System.out.println("Start : processing file  "+infile);
    System.out.println("Start : ecryption used  "+password);
    
    FileOutputStream fos =null;
    try 
    {
      File inputFile = new File(infile);
    
    File outputFile = new File(outfile);
      fos = new FileOutputStream(outputFile);
        Workbook wBook = StreamingReader.builder()   
            .password(password)
            .sstCacheSize(100)    
            .open(new FileInputStream(inputFile));
      // Get first sheet from the workbook
      Sheet sheet = wBook.getSheetAt(0);
      //            Row row;
      Cell cell;
      for(Row row : sheet) 
      {
        if(rowcount==0)
          rownum=row.getLastCellNum();
        rowcount++;
        for(int cn=0; cn<rownum; cn++) 
        {
          // If the cell is missing from the file, generate a blank one
          // (Works by specifying a MissingCellPolicy)
          cell = row.getCell(cn,org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK );
          if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) 
          {
            if(DateUtil.isCellDateFormatted(cell))
            {
              SimpleDateFormat datetemp = new SimpleDateFormat("dd-MM-yyyy");
              try
              {
                String cellvalue = datetemp.format(cell.getDateCellValue());
                if(cn==rownum-1)
                  cellDData.append("\""+cellvalue + "\"");
                else
                  cellDData.append("\""+cellvalue + "\",");  
                continue;
              }
              catch(Exception error)
              {
                msg_log.put("status","unsuccess");
                msg_log.put("error",error.toString());
          System.out.println("Exception in password protected processing file and reading cell value"+error.toString());
                error.printStackTrace();
                return msg_log;
              }
            } 
            if(cn==(rownum-1))cellDData.append("\""+new BigDecimal(cell.getNumericCellValue())+ "\"");
            else cellDData.append("\""+new BigDecimal(cell.getNumericCellValue())+ "\",");
            continue;
          } 
          // Print the cell for debugging
           if(cn==rownum-1)
            cellDData.append("\""+cell.getStringCellValue()+ "\"");
          else 
            cellDData.append("\""+cell.getStringCellValue()+ "\",");
        }
        cellDData.append("\n");
      }
    System.out.println("cellDData");
    //System.out.println("cellDData:"+cellDData.toString());
      fos.write(cellDData.toString().getBytes());
      fos.close();
  
      msg_log.put("status","success");
      msg_log.put("msg","File converted to csv");
      msg_log.put("outFile",outfile);
    } 
    catch (Exception ioe) 
    {
      msg_log.put("status","unsuccess");
      msg_log.put("error",ioe.toString());
    System.out.println("Exception "+ioe.toString());
    }
    finally
    {
      try{
        if(fos!=null)
        {
          fos.close();
          fos=null;
        }
      }
      catch(Exception err){}
    }
    return msg_log;
  }

public HashMap convertXlsxToCSV(String infile, String outfile) 
{
  int rownum=0;
  int rowcount=0;
  HashMap msg_log=new HashMap();
  StringBuffer cellDData = new StringBuffer();
  
  
  FileOutputStream fos =null;
  try 
  {
  	File inputFile = new File(infile);

  
	File outputFile = new File(outfile);


    fos = new FileOutputStream(outputFile);

 
      Workbook wBook = StreamingReader.builder()   
          .sstCacheSize(100)    
          .open(new FileInputStream(inputFile));
    // Get first sheet from the workbook
    Sheet sheet = wBook.getSheetAt(0);
    //            Row row;
    Cell cell;
    for(Row row : sheet) 
    {
      if(rowcount==0)
        rownum=row.getLastCellNum();
      rowcount++;
      
      //System.out.println("rownum : "+rownum);
      for(int cn=0; cn<rownum; cn++) 
      {
        // If the cell is missing from the file, generate a blank one
        // (Works by specifying a MissingCellPolicy)
        cell = row.getCell(cn,org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK );
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) 
        {
          if(DateUtil.isCellDateFormatted(cell))
          {
            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MM-yyyy");
            try
            {

              String cellvalue = datetemp.format(cell.getDateCellValue());
              if(cn==rownum-1)
                cellDData.append("\""+cellvalue + "\"");
              else
                cellDData.append("\""+cellvalue + "\",");  
              continue;
            }
            catch(Exception error)
            {
              msg_log.put("status","unsuccess");
              msg_log.put("error",error.toString());
        System.out.println("Exception in processing file and reading cell value"+error.toString());
              error.printStackTrace();
              return msg_log;
            }
          } 
          if(cn==(rownum-1))cellDData.append("\""+new BigDecimal(cell.getNumericCellValue())+ "\"");
          else cellDData.append("\""+new BigDecimal(cell.getNumericCellValue())+ "\",");
          continue;
        } 
        // Print the cell for debugging
         if(cn==rownum-1)
          cellDData.append("\""+cell.getStringCellValue()+ "\"");
        else 
          cellDData.append("\""+cell.getStringCellValue()+ "\",");
      }
      cellDData.append("\n");
    }
	System.out.println("cellDData");
	//System.out.println("cellDData:"+cellDData.toString());
    fos.write(cellDData.toString().getBytes());
    fos.close();

    msg_log.put("status","success");
    msg_log.put("msg","File converted to csv");
    msg_log.put("outFile",outfile);
  } 
  catch (Exception ioe) 
  {
    msg_log.put("status","unsuccess");
    msg_log.put("error",ioe.toString());
	System.out.println("Exception "+ioe.toString());
  }
  finally
  {
    try{
      if(fos!=null)
      {
        fos.close();
        fos=null;
      }
    }
    catch(Exception err){}
  }
  return msg_log;
}
 
  // public static void main(String[] args) 
  // {
    
  //   ExcelToCSV e=new ExcelToCSV();
  //   System.out.println("Processing XLS File");
  //   HashMap m=e.convertXlsxToCSV("excel.xls", "excel.csv");
  //   System.out.println("Processing XLSX File");
  //   m=e.convertXlsxToCSV("DivCashFlowReg26062018.xlsx", "DivCashFlowReg26062018.csv");
  // }
}

