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
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
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

public class DiscussDataToolForAstraZeneca2018Third {
	public static void main(String[] args) {
		try {
			String rfp_id = "AstraZeneca_2018";
			String location = "RFP";
			  int firstPriceIndex = 87;
//			  int commentIndex = ;
			String sourceFile = "D:/RFP_data/"+rfp_id+"_response.xlsx";
//			String destFile = "D:/RFP_EXCEL_HOTEL/"+rfp_id+"_"+DateUtils.getDate("yyyy-MM-dd")+"议价.xlsx";
			String destFile = "D:/RFP_data/RFP_DATA_EXPORT/"+rfp_id+"_response_"+DateUtils.getDate("yyyy-MM-dd")+".xlsx";
			//int x = 78; 未使用到
			Map<String, Integer> rowNumAndHotelIdMap = new HashMap<String, Integer>();
			/**
			 * step 1 查询历史记录。
			 */
			
//			String sqlForHotelStatus = "SELECT CUSTOMID,DISCUSS_FLAG,SUBMIT_FLAG " 
//													+" FROM T_RFP_CUSTOM_RIGHT WHERE RFPID = '"+rfp_id+"'"
//													+"  AND CUSTOMID NOT IN ('W_TEST') order by CAST( SUBSTRING(CUSTOMID,2,len(CUSTOMID)) AS int)";
			String sqlForHotelStatus = "SELECT r.CUSTOMID,r.DISCUSS_FLAG,r.SUBMIT_FLAG ,r.HOTEL_RFP_STATUS,c.NAME_CN, r.UPDATETIME,ISNULL(g.GSO,'N/A') AS GSO" 
				+" FROM T_RFP_CUSTOM_RIGHT  r inner join  T_RFP_CUSTOM c "
				+" on r.CUSTOMID = c.CUSTOM_ID"
				+" left join T_RFP_GSO_HOTEL g"
				+" on g.HOTEL_ID = c.CUSTOM_ID"
				+" WHERE r.RFPID = '"+rfp_id+"'"
//				+" AND r.DISCUSS_FLAG = 2  AND r.SUBMIT_FLAG = 1 "
				+" AND r.CUSTOMID  not in ('AZ29','AZ96','AZ181','AZ233','AZ296','AZ301', 'AZ320','AZ32','AZ336')"
//				+" AND r.CUSTOMID = 'A222'"
				//+" AND r.HOTEL_RFP_STATUS = 1 "
				+" order by CAST( SUBSTRING(r.CUSTOMID,3,len(CUSTOMID)) AS int)";
			List listHotelStatusList = DBUtil.querySqlForRfp(sqlForHotelStatus, location);
			if(listHotelStatusList == null || listHotelStatusList.size() == 0){
				System.out.println("RFP : "+rfp_id+" 无数据");
				return ;
			}
			
			FileUtils.copyFile(sourceFile, destFile);
			FileInputStream fis = new FileInputStream(destFile);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			OutputStream out = new FileOutputStream(destFile);
			XSSFSheet sheet = workbook.getSheet("AZ_DM");
			
			XSSFFont font = workbook.createFont(); 
			font.setFontName("Arial");
			font.setFontHeightInPoints((short)10);

			XSSFFont fontWithColor = workbook.createFont();
			fontWithColor.setFontName("Arial");
			fontWithColor.setFontHeightInPoints((short)10);
			//fontWithColor.setColor(HSSFColor.RED.index);
			XSSFFont fontWithColor2 = workbook.createFont();
			fontWithColor2.setFontName("Arial");
			fontWithColor2.setFontHeightInPoints((short)10);
			fontWithColor2.setColor(HSSFColor.BLUE.index);
			//
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFont(font);
			cellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			//todo
			XSSFDataFormat df = workbook.createDataFormat();
			
			cellStyle.setDataFormat(df.getFormat("#,#0.00"));
			//todo
			//文本对齐
			CellStyle cellAlign = workbook.createCellStyle();
			cellAlign.setFont(font);
		//	cellAlign.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			cellAlign.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			//文本对齐 右
			CellStyle cellAlignRight = workbook.createCellStyle();
			cellAlign.setFont(font);
		  //cellAlign.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			cellAlign.setDataFormat(df.getFormat("#,#0.00"));
			cellAlign.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			//颜色字体
			CellStyle cellColor =workbook.createCellStyle();
			cellColor.setFont(fontWithColor);
			//颜色字体2
			CellStyle cellColor2 =workbook.createCellStyle();
			cellColor2.setFont(fontWithColor2);
			
			
			
			String sqlForHisData = "WITH TEMP_TAB AS ( "
				+	"select *, ROW_NUMBER() OVER( partition by  QUESTION_ID , HOTEL_ID  ORDER BY CREATEDATATIME  DESC ) AS ROWNUM"    
				+"   from T_RFP_HOTEL_HISTORY "
				+" where RFP_ID = '"+rfp_id+"' " 
				+" AND  HOTEL_ID NOT IN ('AZ29','AZ96','AZ181','AZ233','AZ296','AZ301', 'AZ320','AZ32','AZ336')"
				+")"
				+" SELECT HOTEL_ID, QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  QUESTION_ORDER_ID "  
				+" FROM  TEMP_TAB WHERE ROWNUM = 1 "
				+" ORDER BY CAST( SUBSTRING(HOTEL_ID,3,len(HOTEL_ID)) AS int) ,  QUESTION_ORDER_ID ";
			List listHotelAllDataList = DBUtil.querySqlForRfp(sqlForHisData, location);
			
			for (int i = 0; i < listHotelStatusList.size(); i++) {
					XSSFRow row = sheet.createRow(i+2);
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
					String gso= (String) map.get("GSO");
					
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
					}
					else if("2".equalsIgnoreCase(submitFlag) && "1".equalsIgnoreCase(discussFlag)){
						status = "第二轮已保存";
					}else if("-1".equalsIgnoreCase(submitFlag) && "2".equalsIgnoreCase(discussFlag)){
						status = "第三轮未提交";
					}else if("1".equalsIgnoreCase(submitFlag) && "2".equalsIgnoreCase(discussFlag)){
						status = "第三轮已提交";
					}
					else if("2".equalsIgnoreCase(submitFlag) && "2".equalsIgnoreCase(discussFlag)){
						
						status = "第三轮已保存";
					}
					
					String rfpStatus = "未知";
					if("0".equalsIgnoreCase(hotelRfpStatus) && "0".equalsIgnoreCase(discussFlag) || "0".equalsIgnoreCase(hotelRfpStatus) && "1".equalsIgnoreCase(discussFlag)){
						rfpStatus = "未锁住";
					}
					else if("1".equalsIgnoreCase(hotelRfpStatus) && "0".equalsIgnoreCase(discussFlag)||("1".equalsIgnoreCase(discussFlag))&&"1".equalsIgnoreCase(hotelRfpStatus)){
						rfpStatus = "已锁住";
					}
				
					
					
					
				
					System.out.println(cellAlign.getFontIndex());
					System.out.println(cellColor.getFontIndex());
					XSSFCell cell = row.createCell(0);
					cell.setCellValue(hotelID);
					cell.setCellStyle(cellStyle);
					cell = row.createCell(1);
					cell.setCellValue(status);
					cell.setCellStyle(cellStyle);
					
					cell = row.createCell(2);
					cell.setCellValue(nameCn);
					cell.setCellStyle(cellAlign);
					
					cell = row.createCell(3);
					cell.setCellValue(updateTime);
					cell.setCellStyle(cellAlign);
					
					cell = row.createCell(4);
					cell.setCellValue(rfpStatus);
					cell.setCellStyle(cellAlign);
					
					
					cell = row.createCell(5);
					cell.setCellValue(gso);
					cell.setCellStyle(cellAlign);

				 List<LinkedHashMap<String, String>> listHotelDataList = new ArrayList<LinkedHashMap<String, String>>();
				for (int j = 0; j < listHotelAllDataList.size(); j++) {
					 LinkedHashMap<String, String> map2 = (LinkedHashMap<String, String>) listHotelAllDataList.get(j);
					if(hotelID.equalsIgnoreCase(map2.get("HOTEL_ID")) ){
						listHotelDataList.add(map2);
					}
				}
				 for (int j = 0; j < listHotelDataList.size(); j++) {
					 Map<String, String> map2 = (Map<String, String>) listHotelDataList.get(j);
					    String questionId= (String) map2.get("QUESTION_ID");
					    int orderId= Integer.parseInt(map2.get("QUESTION_ORDER_ID"));
						String answer= (String) map2.get("ANSWER");
//						String orderID= (String) map2.get("QUESTION_ORDER_ID");
						XSSFCell cell2 = row.createCell(j+6);		
						cell2.setCellValue(answer);
						cell2.setCellStyle(cellStyle);
						
						//todo
						if(orderId == 30){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
							
						}else if(orderId == 31){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
						}else if(orderId == 34){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
						}else if(orderId == 35){
							if( ! NumberUtils.isNumber(answer)){
								answer = "0";
							}
							cell2.setCellValue(Double.parseDouble(answer));
						}
						
						//todo 
						//以下内容在做第二轮时候再放开注释
						if(30 == orderId){
							 XSSFCell priceCellTemp =  row.createCell(firstPriceIndex);
							 priceCellTemp.setCellValue(Double.parseDouble(answer));
							 priceCellTemp.setCellStyle(cellStyle);
							 //如下内容做第三轮再放开 第三轮target
						 XSSFCell priceCellTemp1 =  row.createCell(firstPriceIndex+10);
							 priceCellTemp1.setCellValue(answer);
							 priceCellTemp1.setCellStyle(cellStyle);
						}else if(31 == orderId){
							 XSSFCell priceCellTemp =  row.createCell(firstPriceIndex+1);
							 priceCellTemp.setCellValue(Double.parseDouble(answer));
							 priceCellTemp.setCellStyle(cellStyle);
							 XSSFCell priceCellTemp1 =  row.createCell(firstPriceIndex+11);
							 priceCellTemp1.setCellValue(answer);
							 priceCellTemp1.setCellStyle(cellStyle);
						}else if(34 == orderId){
							 XSSFCell priceCellTemp =  row.createCell(firstPriceIndex+2);
							 priceCellTemp.setCellValue(Double.parseDouble(answer));
							 priceCellTemp.setCellStyle(cellStyle);
							 XSSFCell priceCellTemp1 =  row.createCell(firstPriceIndex+12);
							 priceCellTemp1.setCellValue(answer);
							 priceCellTemp1.setCellStyle(cellStyle);
						}else if(35 == orderId){
							 XSSFCell priceCellTemp =  row.createCell(firstPriceIndex+3);
							 priceCellTemp.setCellValue(Double.parseDouble(answer));
							 priceCellTemp.setCellStyle(cellStyle);
							 XSSFCell priceCellTemp1 =  row.createCell(firstPriceIndex+13);
							 priceCellTemp1.setCellValue(answer);
							 priceCellTemp1.setCellStyle(cellStyle);
						}
				}
				 
				 
				 String sqlForMaxStep  = "select ISNULL(max(DISCUSS_STEP),'0') as DISCUSS_STEP from T_RFP_DISCUSS a where a.HOTEL_ID = '"+hotelID+"' and a.RFP_ID = '"+rfp_id+"'";
				 /**
				  * 查询第二轮报价问题
				  */
					List stepList = DBUtil.querySqlForRfp(sqlForMaxStep, location);	
					if(stepList != null || stepList.size() != 0){
							Map<String, String> stepMap = (Map<String, String>) stepList.get(0);
//							String maxStep = stepMap.get("DISCUSS_STEP");
							int maxStep = Integer.parseInt(stepMap.get("DISCUSS_STEP"));
							/**
							 * 客户方
							 */
							for (int j = 2; j <= maxStep; j++) {
								String sqlForHisDisData = "WITH TEMP_TAB AS ( "
									+	"select *, ROW_NUMBER() OVER( partition by  QUESTION_ID , HOTEL_ID  ORDER BY CREATEON  DESC ) AS ROWNUM"    
									+"   from T_RFP_DISCUSS "
									+" where RFP_ID = '"+rfp_id+"' " 
									+" AND  HOTEL_ID  = '"+hotelID+"'"
									+" AND CREATEUSER = '客户方' "
									+" AND DISCUSS_STEP = '"+j+"'"
									+")"
									+" SELECT HOTEL_ID, QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  ORDER_ID "  
									+" FROM  TEMP_TAB WHERE ROWNUM = 1 "
									+" ORDER BY CAST( SUBSTRING(HOTEL_ID,3,len(HOTEL_ID)) AS int) ,  ORDER_ID ";
								List listHotelAllDataDisList = DBUtil.querySqlForRfp(sqlForHisDisData, location);
								for (int k = 0; k < listHotelAllDataDisList.size(); k++) {
									Map<String, String> dataDisMap = (Map<String, String>) listHotelAllDataDisList.get(k);
									String answer= (String) dataDisMap.get("ANSWER");
									int orderId= Integer.parseInt( dataDisMap.get("ORDER_ID"));
									int destCell = firstPriceIndex+orderId-35+(j-2)*10; //客户方target内容填充  
									if(30 == orderId){
										XSSFCell priceCell = row.getCell(destCell);
//										XSSFCell priceCell = row.getCell(firstPriceIndex-5);
										 if(priceCell == null){
											 priceCell = row.createCell(destCell);
											 priceCell.setCellValue(answer); 
											 priceCell.setCellStyle(cellAlignRight);
										 }else{
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }
									}else if(31 == orderId){
										XSSFCell priceCell = row.getCell(destCell);
//										XSSFCell priceCell = row.getCell(firstPriceIndex-4);
										 if(priceCell == null){
											 priceCell = row.createCell(destCell);
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }else{
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }
									}else if(34 == orderId){
										XSSFCell priceCell = row.getCell(destCell-2);//减去2的原因 TODO
//										XSSFCell priceCell = row.getCell(firstPriceIndex-3);
										 if(priceCell == null){
											 priceCell = row.createCell(destCell-2);
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }else{
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }
									}else if(35 == orderId){
										XSSFCell priceCell = row.getCell(destCell-2);
//										XSSFCell priceCell = row.getCell(firstPriceIndex-2);
										 if(priceCell == null){
											 priceCell = row.createCell(destCell-2);
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }else{
											 priceCell.setCellValue(answer);
											 priceCell.setCellStyle(cellAlignRight);
										 }
									}
//									 if(maxStep != j){
//											XSSFCell xssfCell = null;
//											if(orderId == 30 || orderId == 31){
//												 xssfCell = row.createCell(destCell+(j-2)*10);
//											}else if(orderId == 34 || orderId == 35){
//												 xssfCell = row.createCell(destCell+(j-2)*10-2);
//											}
//											xssfCell.setCellValue(answer);
//											xssfCell.setCellStyle(cellAlign);											
//											
//										}
								}
							}
							/**
							 * 酒店方
							 */
							for (int j = 2; j <= maxStep; j++) {
								String sqlForHisDisData = "WITH TEMP_TAB AS ( "
									+	"select *, ROW_NUMBER() OVER( partition by  QUESTION_ID , HOTEL_ID  ORDER BY CREATEON  DESC ) AS ROWNUM"    
									+"   from T_RFP_DISCUSS "
									+" where RFP_ID = '"+rfp_id+"' " 
									+" AND  HOTEL_ID  = '"+hotelID+"'"
									+" AND CREATEUSER = '酒店方' "
									+" AND DISCUSS_STEP = '"+j+"'"
									+")"
									+" SELECT HOTEL_ID, QUESTION_ID , ISNULL(ANSWER,'') AS ANSWER ,  ORDER_ID "  
									+" FROM  TEMP_TAB WHERE ROWNUM = 1 "
									+" ORDER BY CAST( SUBSTRING(HOTEL_ID,3,len(HOTEL_ID)) AS int) ,  ORDER_ID ";
								List listHotelAllDataDisList = DBUtil.querySqlForRfp(sqlForHisDisData, location);
								for (int k = 0; k < listHotelAllDataDisList.size(); k++) {
									Map<String, String> dataDisMap = (Map<String, String>) listHotelAllDataDisList.get(k);
									String answer= (String) dataDisMap.get("ANSWER");
									 Double answerNum = Double.parseDouble(answer);
									int orderId= Integer.parseInt( dataDisMap.get("ORDER_ID"));
									 String hotelId = dataDisMap.get("HOTEL_ID");
									int destCell = firstPriceIndex+orderId-30+(j-2)*10;
									if(30 == orderId){
										XSSFCell priceCell = row.getCell(destCell);
										
										 String sql = "select top 1 a.ANSWER from  T_RFP_HOTEL_HISTORY a where  a.HOTEL_ID = '"+hotelId+"' and a.RFP_ID = 'AstraZeneca_2018' AND QUESTION_ID = '30'  order by   a.CREATEDATATIME desc";
									     List data = DBUtil.querySqlForRfp(sql, location);
									     LinkedHashMap<String, String> answers = (LinkedHashMap<String, String>) data.get(0);
										 String firstAnswer = answers.get("ANSWER");
									     System.out.println("---@@@@@@---"+firstAnswer);
									     Double answerss = Double.parseDouble(firstAnswer);
										//此处有条记录AZ98,需要先改成单个价格的，再改回去。
									     if(answerNum < answerss){
									    	 XSSFFont fontWithColors = workbook.createFont();
												fontWithColor.setFontName("Arial");
												fontWithColor.setFontHeightInPoints((short)10);
												CellStyle cellColors =workbook.createCellStyle();
												cellColors.setDataFormat(df.getFormat("#,#0.00"));
												//cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());//
												if(j == 2){
													cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());//	
												}
												else{
													cellColors.setFillForegroundColor(IndexedColors.WHITE.getIndex());//	fontWithColor2.setColor(HSSFColor.BLUE.index);
												}
												cellColors.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cellColors.setFont(fontWithColors);
												priceCell.setCellValue(answerNum);
												priceCell.setCellStyle(cellColors);
									    	 
									     }else{
									    	 
									    		CellStyle CellStyle =workbook.createCellStyle();
											    CellStyle.setDataFormat(df.getFormat("#,#0.00"));
											    priceCell.setCellValue(answerNum);
											    priceCell.setCellStyle(cellStyle);
									    	 
									     }
									     
										
										
									
									}else if(31 == orderId){
										XSSFCell priceCell = row.getCell(destCell);
										String sql = "select top 1 a.ANSWER from  T_RFP_HOTEL_HISTORY a where  a.HOTEL_ID = '"+hotelId+"' and a.RFP_ID = 'AstraZeneca_2018' AND QUESTION_ID = '31'  order by   a.CREATEDATATIME desc";
									     List data = DBUtil.querySqlForRfp(sql, location);
									     LinkedHashMap<String, String> answers = (LinkedHashMap<String, String>) data.get(0);
										 String firstAnswer = answers.get("ANSWER");
									     System.out.println("---@@@@@@---"+firstAnswer);
									     Double answerss = Double.parseDouble(firstAnswer);
										
									     if(answerNum < answerss){
									    	 XSSFFont fontWithColors = workbook.createFont();
												fontWithColor.setFontName("Arial");
												fontWithColor.setFontHeightInPoints((short)10);
												CellStyle cellColors =workbook.createCellStyle();
												cellColors.setDataFormat(df.getFormat("#,#0.00"));
												//cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
												if(j == 2){
													cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());//	
												}
												else{
													cellColors.setFillForegroundColor(IndexedColors.WHITE.getIndex());//
												}
												cellColors.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cellColors.setFont(fontWithColors);
												priceCell.setCellValue(answerNum);
												priceCell.setCellStyle(cellColors);
									    	 
									     }else{
									    	 
									    		CellStyle CellStyle =workbook.createCellStyle();
											    CellStyle.setDataFormat(df.getFormat("#,#0.00"));
											    priceCell.setCellValue(answerNum);
											    priceCell.setCellStyle(cellStyle);
									    	 
									     }
										
		
									}else if(34 == orderId){
										XSSFCell priceCell = row.getCell(destCell-2);
										
										String sql = "select top 1 a.ANSWER from  T_RFP_HOTEL_HISTORY a where  a.HOTEL_ID = '"+hotelId+"' and a.RFP_ID = 'AstraZeneca_2018' AND QUESTION_ID = '34'  order by   a.CREATEDATATIME desc";
									     List data = DBUtil.querySqlForRfp(sql, location);
									     LinkedHashMap<String, String> answers = (LinkedHashMap<String, String>) data.get(0);
										 String firstAnswer = answers.get("ANSWER");
									     System.out.println("---@@@@@@---"+firstAnswer);
									     Double answerss = Double.parseDouble(firstAnswer);
										
									     if(answerNum < answerss){
									    	 XSSFFont fontWithColors = workbook.createFont();
												fontWithColor.setFontName("Arial");
												fontWithColor.setFontHeightInPoints((short)10);
												CellStyle cellColors =workbook.createCellStyle();
												cellColors.setDataFormat(df.getFormat("#,#0.00"));
												//cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
												if(j == 2){
													cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());//	
												}
												else{
													cellColors.setFillForegroundColor(IndexedColors.WHITE.getIndex());//
												}
												cellColors.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cellColors.setFont(fontWithColors);
												priceCell.setCellValue(answerNum);
												priceCell.setCellStyle(cellColors);
									    	 
									     }else{
									    	 
									    		CellStyle CellStyle =workbook.createCellStyle();
											    CellStyle.setDataFormat(df.getFormat("#,#0.00"));
											    priceCell.setCellValue(answerNum);
											    priceCell.setCellStyle(cellStyle);
									    	 
									     }
	
									}else if(35 == orderId){
										XSSFCell priceCell = row.getCell(destCell-2);
										
										String sql = "select top 1 a.ANSWER from  T_RFP_HOTEL_HISTORY a where  a.HOTEL_ID = '"+hotelId+"' and a.RFP_ID = 'AstraZeneca_2018' AND QUESTION_ID = '35'  order by   a.CREATEDATATIME desc";
									     List data = DBUtil.querySqlForRfp(sql, location);
									     LinkedHashMap<String, String> answers = (LinkedHashMap<String, String>) data.get(0);
										 String firstAnswer = answers.get("ANSWER");
									     System.out.println("---@@@@@@---"+firstAnswer);
									     Double answerss = Double.parseDouble(firstAnswer);
										
									     if(answerNum < answerss){
									    	 XSSFFont fontWithColors = workbook.createFont();
												fontWithColor.setFontName("Arial");
												fontWithColor.setFontHeightInPoints((short)10);
												CellStyle cellColors =workbook.createCellStyle();
												cellColors.setDataFormat(df.getFormat("#,#0.00"));
												//cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
												if(j == 2){
													cellColors.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());//	
												}
												else{
													cellColors.setFillForegroundColor(IndexedColors.WHITE.getIndex());//
												}
												cellColors.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cellColors.setFont(fontWithColors);
												priceCell.setCellValue(answerNum);
												priceCell.setCellStyle(cellColors);
									    	 
									     }else{
									    	 
									    		CellStyle CellStyle =workbook.createCellStyle();
											    CellStyle.setDataFormat(df.getFormat("#,#0.00"));
											    priceCell.setCellValue(answerNum);
											    priceCell.setCellStyle(cellStyle);
									    	 
									     }
										

									}
									/**
									 * if have next step, extra data
									 */
									if( 2==j){
										if(orderId == 30 || orderId == 31){
											XSSFCell xssfCell = row.getCell(destCell+(j-1)*10);
											if(xssfCell == null){
												xssfCell = row.createCell(destCell+(j-1)*10);
											}
											xssfCell.setCellValue(answer);
											xssfCell.setCellStyle(cellColor);											
										}else if(orderId == 34 || orderId == 35){
											XSSFCell xssfCell = row.getCell(destCell+(j-1)*10-2);
											if(xssfCell == null){
												xssfCell = row.createCell(destCell+(j-1)*10-2);
											}
											xssfCell.setCellValue(answer);
											xssfCell.setCellStyle(cellColor);
										}
										
									}
								}
							}
							/**
							 * 查询comments
							 *
							 */
							for (int k = 2; k <= maxStep; k++) {
								String sql = "select top 1  *  from T_RFP_DISCUSS_COMMENT where  RFPID = '"+rfp_id+"' and HOTEL_ID = '"+hotelID+"'  and  STEP = '"+k+"' ORDER BY CREATEON DESC";	
								List listComment = DBUtil.querySqlForRfp(sql, location);
								if(listComment != null && listComment.size() != 0){
									Map<String, String> dataCommentMap = (Map<String, String>) listComment.get(0);
									String comment= (String) dataCommentMap.get("COMMENTSALL");
									XSSFCell commentCell = row.createCell(firstPriceIndex+4+(k-2)*10);
									if( !"A119".equalsIgnoreCase(hotelID) ){
										
										commentCell.setCellValue(comment);
									}
									commentCell.setCellStyle(cellAlign);
								}
								
							}
					}
					
			
		
			
			}

			workbook.write(out);
			SendMail mail = new SendMail();
			List<File> fileList = new ArrayList<File>();
			fileList.add(new File(destFile));
			String subject = rfp_id+"_response_"+DateUtils.getDate("yyyy-MM-dd")+"_DM ";
			String body= "阿斯利康DM见附件";
			String sender = ServiceConfig.getProperty("MAIL.FROM");
			String to = "harlan.hou@citsgbt.com";
			//String to1 = "hotelrfpsupport.cn@citsgbt.com";
			// mail.sendMutiMail2AndCC(subject, body, fileList, to,sender);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
//			System.out.println();
		}
		
		
	}
	
}
