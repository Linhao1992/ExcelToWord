package linhao;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;



public class Model {
	
	//private Logger logger = LoggerFactory.getLogger(Model.class);  

    public HSSFWorkbook wb;  
    public HSSFSheet sheet;  
    public HSSFRow row;
    public HSSFCell cell;
    public Integer beginRow,endRow,colNum,rowNum;
    public Integer id,huming,nianyueri,yuefen,lixi;
    public String inputdir,outputdir;
    //==================================================================================================读取Excel
	public void readExcel(String filename,String dirname) throws Exception {//Map<账号，Map<金额，姓名>>
		String ext = filename.substring(filename.lastIndexOf("."));
		try {
			InputStream is = new FileInputStream(filename);
			//================================================================================获取xls
			if (".xls".equals(ext)) {//判断excel文件类型并实例化wb对象
				wb = new HSSFWorkbook(is);
			} else {
				wb = null;
			}
            sheet = wb.getSheetAt(0);// 获取工作表
            rowNum = sheet.getLastRowNum(); // 总行数
            
            huming=1;//户名
            lixi=5;//利息
            nianyueri = 6;//年月日
            yuefen=7;//月份

            beginRow = 1;//开始行
            endRow=47;//结束行
            inputdir = filename.replace("汇总.xls", "");
            outputdir = dirname.replace(".", "")+"\\";

          //================================================================================修改区域
            
          for(id=beginRow;id<endRow;id++){
            	row = sheet.getRow(id);	
                String shuruming = inputdir+id+".doc";
                String shuchuming = outputdir+getValue(row.getCell(huming)).trim()+".doc";

                Map<String, String> map=new HashMap<String, String>();  
                map.put("lixi", getValue(row.getCell(lixi)).trim());  
                map.put("nianyueri", getValue(row.getCell(nianyueri)).trim());  
                map.put("yuefen", getValue(row.getCell(yuefen)).trim());  
                
                updateWord(shuruming,shuchuming,map);
        		
                /*
                JOptionPane.showMessageDialog(null, 
                			"Excel文件："+filename+
                			"\r\n第"+id+"家公司："+getValue(row.getCell(huming))+
                			"\r\n账号："+getValue(row.getCell(zhanghao))+
                			"\r\n余额："+getValue(row.getCell(yue))+"万元"+
                			"\r\n利息："+getValue(row.getCell(lixi))+"元");
                */
            }
            System.exit(0);
	}catch (FileNotFoundException e) {// 找不到文件
		System.exit(0);
	} catch (IOException e) {// 读取异常
		System.exit(0);
	}
}
	//================================================================================修改word的方法
	public void updateWord(String inputname, String outputname, Map<String, String> map) {
		try {
			FileInputStream in = new FileInputStream(new File(inputname));
			HWPFDocument hdt = new HWPFDocument(in);
			Range range = hdt.getRange();
			for (Map.Entry<String, String> entry : map.entrySet()) {// 替换文本内容
				range.replaceText(entry.getKey(), entry.getValue());
			}
			ByteArrayOutputStream ostream = new ByteArrayOutputStream();
			FileOutputStream out = new FileOutputStream(outputname);
			hdt.write(ostream);
			out.write(ostream.toByteArray());// 输出字节流
			ostream.close();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	//================================================================================获取单元格值的方法
	public String getValue(HSSFCell cell) {// 获取单元格的值
		DecimalFormat jine = new DecimalFormat("0.00");// 十进制格式
		switch (cell.getCellType()) {
		case 0: // 数字：金额
			return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
		case 1:// 中文：户名/账号
			return String.valueOf(cell.getStringCellValue()).replaceAll("\r|\n", "");
		case 2:// 2公式
			try {
				return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			} catch (IllegalStateException e) {
				return String.valueOf(jine.format(cell.getRichStringCellValue())).replace(" ", "").replaceAll("\r|\n","");
			}
		case 3:// 3空值
			return "单元格为空";
		case 4:// 4真假
			return String.valueOf(cell.getBooleanCellValue()).replace(" ", "").replaceAll("\r|\n", "");
		case 5:// 5错误
			return "单元格内容错误";
		}
		return "没有匹配值类型";
	}
}
