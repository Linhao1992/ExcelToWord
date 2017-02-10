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
    //==================================================================================================��ȡExcel
	public void readExcel(String filename,String dirname) throws Exception {//Map<�˺ţ�Map<������>>
		String ext = filename.substring(filename.lastIndexOf("."));
		try {
			InputStream is = new FileInputStream(filename);
			//================================================================================��ȡxls
			if (".xls".equals(ext)) {//�ж�excel�ļ����Ͳ�ʵ����wb����
				wb = new HSSFWorkbook(is);
			} else {
				wb = null;
			}
            sheet = wb.getSheetAt(0);// ��ȡ������
            rowNum = sheet.getLastRowNum(); // ������
            
            huming=1;//����
            lixi=5;//��Ϣ
            nianyueri = 6;//������
            yuefen=7;//�·�

            beginRow = 1;//��ʼ��
            endRow=47;//������
            inputdir = filename.replace("����.xls", "");
            outputdir = dirname.replace(".", "")+"\\";

          //================================================================================�޸�����
            
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
                			"Excel�ļ���"+filename+
                			"\r\n��"+id+"�ҹ�˾��"+getValue(row.getCell(huming))+
                			"\r\n�˺ţ�"+getValue(row.getCell(zhanghao))+
                			"\r\n��"+getValue(row.getCell(yue))+"��Ԫ"+
                			"\r\n��Ϣ��"+getValue(row.getCell(lixi))+"Ԫ");
                */
            }
            System.exit(0);
	}catch (FileNotFoundException e) {// �Ҳ����ļ�
		System.exit(0);
	} catch (IOException e) {// ��ȡ�쳣
		System.exit(0);
	}
}
	//================================================================================�޸�word�ķ���
	public void updateWord(String inputname, String outputname, Map<String, String> map) {
		try {
			FileInputStream in = new FileInputStream(new File(inputname));
			HWPFDocument hdt = new HWPFDocument(in);
			Range range = hdt.getRange();
			for (Map.Entry<String, String> entry : map.entrySet()) {// �滻�ı�����
				range.replaceText(entry.getKey(), entry.getValue());
			}
			ByteArrayOutputStream ostream = new ByteArrayOutputStream();
			FileOutputStream out = new FileOutputStream(outputname);
			hdt.write(ostream);
			out.write(ostream.toByteArray());// ����ֽ���
			ostream.close();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	//================================================================================��ȡ��Ԫ��ֵ�ķ���
	public String getValue(HSSFCell cell) {// ��ȡ��Ԫ���ֵ
		DecimalFormat jine = new DecimalFormat("0.00");// ʮ���Ƹ�ʽ
		switch (cell.getCellType()) {
		case 0: // ���֣����
			return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
		case 1:// ���ģ�����/�˺�
			return String.valueOf(cell.getStringCellValue()).replaceAll("\r|\n", "");
		case 2:// 2��ʽ
			try {
				return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			} catch (IllegalStateException e) {
				return String.valueOf(jine.format(cell.getRichStringCellValue())).replace(" ", "").replaceAll("\r|\n","");
			}
		case 3:// 3��ֵ
			return "��Ԫ��Ϊ��";
		case 4:// 4���
			return String.valueOf(cell.getBooleanCellValue()).replace(" ", "").replaceAll("\r|\n", "");
		case 5:// 5����
			return "��Ԫ�����ݴ���";
		}
		return "û��ƥ��ֵ����";
	}
}
