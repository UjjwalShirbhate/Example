package ApacheExcel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

// user/programer will create  their data they will then write it in correct column position in a  excel file
public class WriteForEach
{
    public static void main(String[] args) throws IOException {
        //we are going  to create file
        XSSFWorkbook workbook =new XSSFWorkbook() ;
        XSSFSheet sheet=workbook.createSheet("DATA") ;
        Object[][] data={
                {"name","Age","Gender"},
                {"amit",22,"Male"},
                {"nikita",21,"Female"}
        };
int rowCount=0;
for(Object[] rowArr:data)
{
    XSSFRow row=sheet.createRow(rowCount);//create a new at current positon
    int col=0;
    for(Object value:rowArr){//for evrey single value in current row
        XSSFCell cell=row.createCell(col);//create a new box/cell

        if(value instanceof Integer )
            cell.setCellValue((Integer )  value) ;
        if(value instanceof String  )
            cell.setCellValue((String )   value) ;
        if(value instanceof Double  )
            cell.setCellValue((Double )   value) ;
        col++;
    }
    rowCount++;
    }
        FileOutputStream f1=new FileOutputStream("Students.xlsx");
        workbook.write(f1);
        f1.close();
    }
}
