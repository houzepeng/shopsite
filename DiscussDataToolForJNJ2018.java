package rfputils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.core.util.DBUtil;
import org.jeecgframework.core.util.DateUtils;
import org.jeecgframework.core.util.FileUtils;
import org.jeecgframework.core.util.SendMail;
import org.jeecgframework.core.util.ServiceConfig;

public class DiscussDataToolForJNJ2018 {
	@SuppressWarnings("unchecked")
	public static void main(String[] args) {
		try {
			String rfp_id = "Johnson_2018";
			String location = "RFP";
			
			String sourceFile = "D:/RFP_data/"+rfp_id+"_response.xlsx";
			String destFile = "D:/RFP_data/RFP_DATA_EXPORT/"+rfp_id+"_"+DateUtils.getDate("yyyy-MM-dd")+".xlsx";
			int x = 97;
			/**
			 * @rowNumAndHotelIdMap 存放hotelID,rowNum的map
			 */
			Map<String, Integer> rowNumAndHotelIdMap = new HashMap<String, Integer>();
			
			/**
			 * step 1 查询历史记录。
			 */
			
			String sqlForHotelStatus = "SELECT r.CUSTOMID,r.DISCUSS_FLAG,r.SUBMIT_FLAG ,r.HOTEL_RFP_STATUS,r.UPDATETIME,c.NAME_CN,ISNULL(g.GSO,'N/A') AS GSO " 
													+" FROM T_RFP_CUSTOM_RIGHT  r inner join  T_RFP_CUSTOM c "
													+" on r.CUSTOMID = c.CUSTOM_ID left join T_RFP_GSO_HOTEL g on g.HOTEL_ID = c.CUSTOM_ID"
													+" WHERE r.RFPID = '"+rfp_id+"'"
													+"  AND r.CUSTOMID  IN ('JS140') order by CAST( SUBSTRING(r.CUSTOMID,3,len(CUSTOMID)) AS int)";
			/**
			 *  listHotelStatusList ,酒店基本状态list
			 */
			List listHotelStatusList = DBUtil.querySqlForRfp(sqlForHotelStatus, location);
			if(listHotelStatusList == null || listHotelStatusList.size() == 0){
				System.out.println("RFP : "+rfp_id+" 无数据");
				return ;
			}
			
			FileUtils.copyFile(sourceFile, destFile);
			FileInputStream fis = new FileInputStream(destFile);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			OutputStream out = new FileOutputStream(destFile);
			XSSFSheet sheet = workbook.getSheet("Hotel Response");
			
			//10号Arial 字体
			XSSFFont font = workbook.createFont(); 
			font.setFontName("Arial");
			font.setFontHeightInPoints((short)10);
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFont(font);
			XSSFDataFormat df = workbook.createDataFormat();
				
			//cellStyle.setDataFormat(df.getFormat("#,#0.00"));
			
			//文本对齐
			CellStyle cellAlign = cellStyle;
			cellAlign.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			
			String sqlForHisData = "WITH TEMP_TAB AS ( "
				+	"select *, ROW_NUMBER() OVER( partition by  QUESTION_ID , HOTEL_ID  ORDER BY CREATEDATATIME  DESC ) AS ROWNUM"    
				+"   from T_RFP_HOTEL_HISTORY "
				+" where RFP_ID = '"+rfp_id+"'  AND QUESTION_ID NOT IN (46,48) " 
				+" AND  HOTEL_ID NOT IN ('GSO_JS1')"
				+")"
				+" SELECT HOTEL_ID, QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  QUESTION_ORDER_ID "  
				+" FROM  TEMP_TAB WHERE ROWNUM = 1 "
				+" ORDER BY CAST( SUBSTRING(HOTEL_ID,3,len(HOTEL_ID)) AS int) ,  QUESTION_ORDER_ID ";
			//酒店第一轮报价的基础信息list
			List listHotelAllDataList = DBUtil.querySqlForRfp(sqlForHisData, location);
			
			for (int i = 0; i < listHotelStatusList.size(); i++) {
					XSSFRow row = sheet.createRow(i+3);
					int rowNum = row.getRowNum();
					//没用
//					row.setRowStyle(cellStyle);
					Map<String, String> map = (Map<String, String>) listHotelStatusList.get(i);
					String hotelID= (String) map.get("CUSTOMID");
					String discussFlag= (String) map.get("DISCUSS_FLAG");
					String submitFlag= (String) map.get("SUBMIT_FLAG");
					String nameCn= (String) map.get("NAME_CN");
					String updateTime = (String)map.get("UPDATETIME");//更新时间
					String hotelRfpStatus = (String)map.get("HOTEL_RFP_STATUS");//是否锁住
					String gsoMail= (String) map.get("GSO");//邮箱
					//map存放 每个酒店记录的rowNum
					rowNumAndHotelIdMap.put(hotelID, rowNum);

					String status = "未知";
					if("-1".equalsIgnoreCase(submitFlag) && "0".equalsIgnoreCase(discussFlag)){
						status = "第一轮未提交";
					}else if("1".equalsIgnoreCase(submitFlag) && "0".equalsIgnoreCase(discussFlag)){
						status = "第一轮已提交";
					}else if("2".equalsIgnoreCase(submitFlag) && "0".equalsIgnoreCase(discussFlag)){
						status = "第一轮已保存";
					}else if("-1".equalsIgnoreCase(submitFlag) && "1".equalsIgnoreCase(discussFlag)){
						status = "第二轮未提交";
					}else if("1".equalsIgnoreCase(submitFlag) && "1".equalsIgnoreCase(discussFlag)){
						status = "第二轮已提交";
					}else if("2".equalsIgnoreCase(submitFlag) && "1".equalsIgnoreCase(discussFlag)){
						status = "第二轮已保存";
					}else if("-1".equalsIgnoreCase(submitFlag) && "2".equalsIgnoreCase(discussFlag)){
						status = "第三轮未提交";
					}else if("1".equalsIgnoreCase(submitFlag) && "2".equalsIgnoreCase(discussFlag)){
						status = "第三轮已提交";
					}
					
					String rfpStatus = "未知";
					if("0".equalsIgnoreCase(hotelRfpStatus) && "0".equalsIgnoreCase(discussFlag)){
						rfpStatus = "未锁住";
					}
					else{
						rfpStatus = "已锁住";
					}
					
					//创建5个
				
					/*for (int j = 4; j < 109; j++) {
						XSSFCell cells = row.createCell(0);//--->创建一个单元格  
						XSSFRow rows = sheet.getRow(j);
						System.out.println("--"+j);
						XSSFCell  cell0 = rows.getCell(0);// 取得城市
						XSSFCell  cell1 = rows.getCell(1);//取得地区	
					//	System.out.println("--"+city+"--"+district);
					//	System.out.println("--"+city.getStringCellValue());
						String city = cell0.getStringCellValue();
						if(city.equals(null)){
							city = "";
						}
						System.out.println("--"+city);
						String district = cell1.getStringCellValue();
						if(district.equals(null)){
							district = "";
						}
						System.out.println("--"+district);
						//cells = row.createCell(0);
						cells.setCellValue(city);
						cells.setCellStyle(cellStyle);
						
						cells = row.createCell(1);
						cells.setCellValue(district);
						cells.setCellStyle(cellStyle);

					}*/
				
					
					XSSFCell cell = row.createCell(0);//--->创建一个单元格  
					cell.setCellValue(hotelID);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(1);
					cell.setCellValue(status);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(2);
					cell.setCellValue(rfpStatus);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(3);
					cell.setCellValue(updateTime);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(4);
					cell.setCellValue(nameCn);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(5);
					cell.setCellValue(gsoMail);
					cell.setCellStyle(cellStyle);
					
					
					
					
					
								
		
					/**
					 * WITH TEMP_TAB AS (
						select *, ROW_NUMBER()     
							OVER( partition by  QUESTION_ID , HOTEL_ID ORDER BY CREATEDATATIME  DESC ) AS ROWNUM  from T_RFP_HOTEL_HISTORY 
						where RFP_ID = 'Johnson_2017' and HOTEL_ID NOT IN ('J_TEST1','J_TEST2') AND QUESTION_ID NOT IN (46,48) 
						
						)
						SELECT HOTEL_ID, QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  QUESTION_ORDER_ID   
							FROM  TEMP_TAB WHERE ROWNUM = 1
							ORDER BY CAST( SUBSTRING(HOTEL_ID,2,len(HOTEL_ID)) AS int) ,  QUESTION_ORDER_ID 
					 */
//					String sqlForHisData = "WITH TEMP_TAB AS ( "
//														+	"select *, ROW_NUMBER() OVER( partition by  QUESTION_ID ORDER BY CREATEDATATIME  DESC ) AS ROWNUM"    
//														+"   from T_RFP_HOTEL_HISTORY "
//														+" where RFP_ID = '"+rfp_id+"' and HOTEL_ID = '"+hotelID+"' AND QUESTION_ID NOT IN (46,48) " 
//														+" AND  HOTEL_ID NOT IN ('J_TEST1','J_TEST2')"
//														+")"
//														+" SELECT  QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  QUESTION_ORDER_ID "  
//														+" FROM  TEMP_TAB WHERE ROWNUM = 1 "
//														+" ORDER BY QUESTION_ORDER_ID ";
//				 List listHotelDataList = DBUtil.querySqlForRfp(sqlForHisData, location);
				 List<LinkedHashMap<String, String>> listHotelDataList = new ArrayList<LinkedHashMap<String, String>>();
				for (int j = 0; j < listHotelAllDataList.size(); j++) {
					 LinkedHashMap<String, String> map2 = (LinkedHashMap<String, String>) listHotelAllDataList.get(j);
					if(hotelID.equalsIgnoreCase(map2.get("HOTEL_ID")) ){
						listHotelDataList.add(map2);
					}
				}
				 for (int j = 0; j < listHotelDataList.size(); j++) {
					 Map<String, String> map2 = (Map<String, String>) listHotelDataList.get(j);
					    String questionID= (String) map2.get("QUESTION_ID");
						String answer= (String) map2.get("ANSWER");
//						String orderID= (String) map2.get("QUESTION_ORDER_ID");
						XSSFCell cell2 = row.createCell(j+6);
						if("67".equalsIgnoreCase(questionID)){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
							/**
							 * 第一轮价格先放入议价酒店报价部位
							 */
							XSSFCell cellTemp = row.createCell(x+5);
							cellTemp.setCellValue(Double.parseDouble(answer));
							cellTemp.setCellStyle(cellStyle);
							
						}else if("68".equalsIgnoreCase(questionID)){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
							
							XSSFCell cellTemp = row.createCell(x+6);
							cellTemp.setCellValue(Double.parseDouble(answer));
							cellTemp.setCellStyle(cellStyle);
							
						}else{
							cell2.setCellValue(answer);
						}
						cell2.setCellStyle(cellStyle);
				}
					
			}
			
			/**
			 * step 2 查询议价部分
			 */
			
			/**
			 * step 2  first (1)
			 * 查询客户方的内容 
			 * 
			 * select * from T_RFP_DISCUSS where CREATEUSER = '客户方' and RFP_ID = 'Johnson_2017'
			 * JNJ 101开始
			 */
			
			/*String discussSQLCustom = "select  HOTEL_ID,QUESTION_ID,  QUESTION_NAME,ANSWER, COMMMENTS from T_RFP_DISCUSS where CREATEUSER = '客户方' and RFP_ID = '"+rfp_id+"'"
														+" and  HOTEL_ID not in ('J_TEST2','J_TEST1') and  DISCUSS_STEP = 2";
			 List<LinkedHashMap<String, String>> discussCustomListData = DBUtil.querySqlForRfp(discussSQLCustom, location);
			 
//			System.out.println(discussCustomListData);
			 for (int i = 0; i < discussCustomListData.size(); i++) {
				 LinkedHashMap<String, String> mapCustom = (LinkedHashMap<String, String>) discussCustomListData.get(i);
				 String hotelId = mapCustom.get("HOTEL_ID");
				 String qId = mapCustom.get("QUESTION_ID");
//				 String qName = mapCustom.get("QUESTION_NAME");
				 String answer = mapCustom.get("ANSWER");
				 Double answerNum = Double.parseDouble(answer);
				 int rowNum = rowNumAndHotelIdMap.get(hotelId);
				 XSSFRow rowTemp = sheet.getRow(rowNum);
				 if("67".equalsIgnoreCase(qId)){
					 XSSFCell cell1 = rowTemp.createCell(x);
					 cell1.setCellValue(answerNum);
					 cell1.setCellStyle(cellStyle);
				 }else if("68".equalsIgnoreCase(qId)){
					 XSSFCell cell2 = rowTemp.createCell(x+1);
					 cell2.setCellValue(answerNum);
					 cell2.setCellStyle(cellStyle);
				 }
				 
			}*/
			
			
			/**
			 * step 2  second (2)
			 * 查询客户方的内容 
			 */
//			 String sql = "select max(DISCUSS_STEP) from T_RFP_DISCUSS a where a.HOTEL_ID = '"+hotelId+"' and a.RFP_ID = '"+rfp_id+"'";
			 /**
			  *  DISCUSS_STEP = 2
			  */
			 
			 
			 /*String discussSQLHotel = "with HOTEL_TEMP AS( "
				 								+"	SELECT  HOTEL_ID,QUESTION_ID,ANSWER, COMMMENTS  ,ROW_NUMBER() OVER(partition by QUESTION_ID,  HOTEL_ID ORDER BY CREATEON DESC) AS Rownum "
				 								+"			FROM T_RFP_DISCUSS where RFP_ID = 'Johnson_2018' "
				 								+"					and CREATEUSER = '酒店方' "
				 								+"							and  HOTEL_ID not in ('J_TEST2','J_TEST1') "
				 								+"								and  DISCUSS_STEP = 2 "	
		 										+"	) "
												+"	SELECT  HOTEL_ID,QUESTION_ID,ANSWER, COMMMENTS from  HOTEL_TEMP "
												+"	where Rownum = 1";
			 List<LinkedHashMap<String, String>> discussHotelListData = DBUtil.querySqlForRfp(discussSQLHotel, location);
			 System.out.println(discussHotelListData);
			 for (int i = 0; i < discussHotelListData.size(); i++) {
				 LinkedHashMap<String, String> mapHotel = (LinkedHashMap<String, String>) discussHotelListData.get(i);
				 String hotelId = mapHotel.get("HOTEL_ID");
				 String qId = mapHotel.get("QUESTION_ID");
//				 String qName = mapHotel.get("QUESTION_NAME");
				 String answer = mapHotel.get("ANSWER");
				 Double answerNum = Double.parseDouble(answer);
				 String comments = mapHotel.get("COMMMENTS");
				 int rowNum = rowNumAndHotelIdMap.get(hotelId);
				 XSSFRow rowTemp = sheet.getRow(rowNum);
				 if("67".equalsIgnoreCase(qId)){
					 XSSFCell cell1 = rowTemp.createCell(x+3);
					 cell1.setCellValue(answerNum);
					 cell1.setCellStyle(cellStyle);
				 }else if("68".equalsIgnoreCase(qId)){
					 XSSFCell cell2 = rowTemp.createCell(x+4);
					 cell2.setCellValue(answerNum);
					 cell2.setCellStyle(cellStyle);
				 }
				 
				 XSSFCell cellCommet = rowTemp.getCell(x+5);
				 if(cellCommet == null){
					 cellCommet = rowTemp.createCell(x+5);
					 cellCommet.setCellValue(comments);
				 }else{
					 cellCommet.setCellValue(cellCommet.toString()+comments);
				 }
				 cellCommet.setCellStyle(cellAlign);
				 
			 }*/
			 /**
			  *  DISCUSS_STEP =3
			  */
			/* for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
				 	XSSFRow rowTemp = sheet.getRow(i);
				 	XSSFCell  cell3 = rowTemp.getCell(x+3);
				 	XSSFCell  cell5 = rowTemp.getCell(x+4);
				 	
				 	if(cell3!= null && cell3.toString()!=""){
				 		XSSFCell cell4 = rowTemp.getCell(x+9);
				 		if(null  == cell4){
				 			cell4 = rowTemp.createCell(x+9);
				 		}
				 		cell4.setCellValue(Double.parseDouble(cell3.toString()));
				 		cell4.setCellStyle(cellStyle);
				 	}
				 	if(null != cell3 && cell5.toString()!=""){
				 		XSSFCell cell6 = rowTemp.getCell(x+10);
				 		if(null == cell6){
				 			cell6 = rowTemp.createCell(x+10);
				 		}
				 		cell6.setCellValue(cell5.toString());
				 		cell6.setCellStyle(cellStyle);
				 	}
			}*/

				/*String discussSQLCustom3 = "select  HOTEL_ID,QUESTION_ID,  QUESTION_NAME,ANSWER, COMMMENTS from T_RFP_DISCUSS where CREATEUSER = '客户方' and RFP_ID = '"+rfp_id+"'"
															+" and  HOTEL_ID not in ('J_TEST2','J_TEST1') and  DISCUSS_STEP = 3";
				 List<LinkedHashMap<String, String>> discussCustomListData3 = DBUtil.querySqlForRfp(discussSQLCustom3, location);
				 
//				System.out.println(discussCustomListData);
				 for (int i = 0; i < discussCustomListData3.size(); i++) {
					 LinkedHashMap<String, String> mapCustom = (LinkedHashMap<String, String>) discussCustomListData3.get(i);
					 String hotelId = mapCustom.get("HOTEL_ID");
					 String qId = mapCustom.get("QUESTION_ID");
//					 String qName = mapCustom.get("QUESTION_NAME");
					 String answer = mapCustom.get("ANSWER");
//					 Double answerNum = Double.parseDouble(answer);
					 String answerNum = answer;
					 int rowNum = rowNumAndHotelIdMap.get(hotelId);
					 XSSFRow rowTemp = sheet.getRow(rowNum);
					 if("67".equalsIgnoreCase(qId)){
						 XSSFCell cell1 = rowTemp.createCell(x+6);
						 cell1.setCellValue(answerNum);
						 cell1.setCellStyle(cellStyle);
					 }else if("68".equalsIgnoreCase(qId)){
						 XSSFCell cell2 = rowTemp.createCell(x+7);
						 cell2.setCellValue(answerNum);
						 cell2.setCellStyle(cellStyle);
					 }
					 
				}
			 
			 String discussSQLHotel3 = "with HOTEL_TEMP AS( "
				 +"	SELECT  HOTEL_ID,QUESTION_ID,ANSWER, COMMMENTS  ,ROW_NUMBER() OVER(partition by QUESTION_ID,  HOTEL_ID ORDER BY CREATEON DESC) AS Rownum "
				 +"			FROM T_RFP_DISCUSS where RFP_ID = 'Johnson_2018' "
				 +"					and CREATEUSER = '酒店方' "
				 +"							and  HOTEL_ID not in ('J_TEST2','J_TEST1') "
				 +"								and  DISCUSS_STEP = 3 "	
				 +"	) "
				 +"	SELECT  HOTEL_ID,QUESTION_ID,ANSWER, COMMMENTS from  HOTEL_TEMP "
				 +"	where Rownum = 1";
			 List<LinkedHashMap<String, String>> discussHotelListData3 = DBUtil.querySqlForRfp(discussSQLHotel3, location);
			 System.out.println(discussHotelListData3);
			 for (int i = 0; i < discussHotelListData3.size(); i++) {
				 LinkedHashMap<String, String> mapHotel = (LinkedHashMap<String, String>) discussHotelListData3.get(i);
				 String hotelId = mapHotel.get("HOTEL_ID");
				 String qId = mapHotel.get("QUESTION_ID");
//				 String qName = mapHotel.get("QUESTION_NAME");
				 String answer = mapHotel.get("ANSWER");
				 Double answerNum = Double.parseDouble(answer);
				 String comments = mapHotel.get("COMMMENTS");
				 int rowNum = rowNumAndHotelIdMap.get(hotelId);
				 XSSFRow rowTemp = sheet.getRow(rowNum);
				 if("67".equalsIgnoreCase(qId)){
					 XSSFCell cell1 = rowTemp.createCell(x+9);
					 cell1.setCellValue(answerNum);
					 cell1.setCellStyle(cellStyle);
				 }else if("68".equalsIgnoreCase(qId)){
					 XSSFCell cell2 = rowTemp.createCell(x+10);
					 cell2.setCellValue(answerNum);
					 cell2.setCellStyle(cellStyle);
				 }
				 
				 XSSFCell cellCommet = rowTemp.getCell(x+11);
				 if(cellCommet == null){
					 cellCommet = rowTemp.createCell(x+11);
					 cellCommet.setCellValue(comments);
				 }else{
					 cellCommet.setCellValue(cellCommet.toString()+comments);
				 }
				 cellCommet.setCellStyle(cellAlign);
			 }
			*/
			
			
	
			workbook.write(out);
			/**
			 * 发送邮件
			 */
			SendMail mail = new SendMail();
			List<File> fileList = new ArrayList<File>();
			fileList.add(new File(destFile));
			String subject = rfp_id+" DM";
			String body= "";
			
		//	mail.sendMutiMailWithCC(subject, body, fileList , mailto, sender ,cc);
			
			
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public  static void done(){
		System.out.println("do");
	}
	
	
	public static void  discuss(String rfp_id){
		try {
//			String rfp_id = "Medtronic_2017";
			String location = "RFP";
			//1.查询哪个RFP，酒店集合
			String sql = "select HOTEL_ID from T_RFP_DISCUSS  " 
						+" where RFP_ID = '"+rfp_id+"' and CREATEUSER = '客户方' group by HOTEL_ID";
			/**
			 * 酒店list
			 */
			List listHotel =  DBUtil.querySqlForRfp(sql,location);
			//2.查询酒店数据
			if(listHotel == null || listHotel.size() == 0){
				System.out.println("RFP discuss: "+rfp_id+" 无数据");
				return ;
			}
			FileUtils.copyFile("D:/discuss.xlsx", "D:/RFP_EXCEL_HOTEL/"+rfp_id+"_"+DateUtils.getDate("yyyy-MM-dd")+"议价.xlsx");
			FileInputStream fis = new FileInputStream("D:/discuss.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			OutputStream out = new FileOutputStream("D:/RFP_EXCEL_HOTEL/"+rfp_id+"_"+DateUtils.getDate("yyyy-MM-dd")+"议价.xlsx");
			XSSFSheet sheet = workbook.getSheet("sheet1");
			XSSFRow row; 
			String cell;
			Map<String, Integer> rowNumAndHotelIdMap = new HashMap<String, Integer>();
			
			for (int m = 2; m < sheet.getPhysicalNumberOfRows()&& m < 80; m++) {
				
				row = sheet.getRow(m);
				cell = row.getCell(0).toString();
				rowNumAndHotelIdMap.put(cell, m);
			}
			System.out.println(rowNumAndHotelIdMap);
			
			
			
			
			
			XSSFFont font = workbook.createFont(); 
			font.setFontName("Arial");
			font.setFontHeightInPoints((short)10);
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFont(font);
			
			
			
			
			for (int i = 0; i < listHotel.size(); i++) {
				Map map = (Map) listHotel.get(i);
				String hotelID= (String) map.get("HOTEL_ID");
				if( "H51".equalsIgnoreCase(hotelID)){
					continue;
				}
				int rowNum = rowNumAndHotelIdMap.get(hotelID);
				System.out.println(hotelID);
				String sqlForHotelQuestion = "select QUESTION_ID from T_RFP_DISCUSS " 
						  +" where HOTEL_ID = '"+hotelID+"' and CREATEUSER = '酒店方'  GROUP BY QUESTION_ID order by CAST(QUESTION_ID AS int) ";
				/**
				 * 酒店回复的问题list
				 */
				List listHotelQuestion =  DBUtil.querySqlForRfp(sqlForHotelQuestion,location);
				if(listHotelQuestion == null && listHotelQuestion.size()==0){
					continue;
				}
				String allComments = "";
				int temp = Integer.parseInt(ServiceConfig.getRfpProperty(rfp_id+"_colNum"));
				for (int j = 0; j < listHotelQuestion.size(); j++) {
					Map questionMap = (Map) listHotelQuestion.get(j);
					String questionID = (String) questionMap.get("QUESTION_ID");
					String sqlForREC = " select top 1 * from T_RFP_DISCUSS "
										+" where QUESTION_ID = '"+questionID+"' and HOTEL_ID = '"+hotelID+"' AND CREATEUSER = '酒店方' " 
										+" order by CREATEON desc";
					List top_REC_hotel_list =  DBUtil.querySqlForRfp(sqlForREC,location);
					Map top_REC_hotel_map = (Map) top_REC_hotel_list.get(0);
					String answer = (String) top_REC_hotel_map.get("ANSWER");
					String comment = (String) top_REC_hotel_map.get("COMMMENTS")=="null"?"":(String) top_REC_hotel_map.get("COMMMENTS");
					allComments = allComments + comment;
					
					XSSFRow row_id = sheet.getRow(rowNum);
					XSSFCell cell_id = row_id.createCell(temp);
					temp++;
					cell_id.setCellValue(answer);
					
					cell_id.setCellStyle(cellStyle);
					if(j == listHotelQuestion.size()-1){
						XSSFCell cell_id_comments = row_id.createCell(Integer.parseInt(ServiceConfig.getRfpProperty(rfp_id+"_colNum"))+4);
						cell_id_comments.setCellValue(allComments);
						cell_id_comments.setCellStyle(cellStyle);
					}
				}
				
				
			}
			workbook.write(out);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
		
	}
}
