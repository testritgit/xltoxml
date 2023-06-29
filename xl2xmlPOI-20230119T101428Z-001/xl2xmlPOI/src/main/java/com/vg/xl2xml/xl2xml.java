/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.vg.xl2xml;

/**
 *
 * @author giri
 */
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

public class xl2xml {

    private static final String FILE_PATH = "C:\\Users\\giri\\Downloads\\";
    private static final String FILE_NAME = "QB Model";
    private static final String EXCEL_FILE_LOCATION = FILE_PATH + FILE_NAME + ".xlsx";
    private static final String XML_FILE_LOCATION = FILE_PATH + FILE_NAME + ".xml";
    // private static final String EXCEL_FILE_LOCATION = "C:\\Users\\giri\\Downloads\\Modelpaper1.xlsx";
//    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\giri\\Downloads\\Model paper.xlsx";
    private static final String NS_URL = "https://appconltd.com/banking";
    // private static final String XML_ROOT = "Nursing";
    private static final String XML_ROOT = "Gate";
    //private static final String SubjectName = "Anatomy";
    private static final String SubjectName = "GateQB";
    private static final String PIC_FILE_LOCATION = "C:\\Users\\giri\\Downloads\\PIC1\\";
    private static int[][] iRC = new int[250][20];
    static String cell0S[] = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"};
    static String sheetname;

    public static void main(String[] args) throws ParserConfigurationException, TransformerConfigurationException, TransformerException, FileNotFoundException, IOException {
        FileInputStream excelFile = new FileInputStream(new File(EXCEL_FILE_LOCATION));
        Workbook workbook = null;

//        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
//        DocumentBuilder dBuilder;
//
//        dBuilder = dbFactory.newDocumentBuilder();
//        Document doc = dBuilder.newDocument();
//        //add elements to Document
//        Element rootElement = doc.createElementNS(NS_URL, XML_ROOT);

        class PIC {

            private int row;
            private int col;
            private String picName;

            // constructor
            public PIC(int row, int col, String picName) {
                this.row = row;
                this.col = col;
                this.picName = picName;
            }

            // getter
            public int getRow() {
                return row;
            }

            public int getCol() {
                return col;
            }

            public String getPicName() {
                return picName;
            }

            // setter
            public void setRow(int row) {
                this.row = row;
            }

            public void setCol(int col) {
                this.col = col;
            }

            public void setPICName(String picName) {
                this.picName = picName;
            }

        }

        List<PIC> listPIC = new LinkedList<PIC>();

        try {
            if (EXCEL_FILE_LOCATION.split("\\.")[1].equalsIgnoreCase("xls")) {
                workbook = new HSSFWorkbook(excelFile);
            } else {
                workbook = new XSSFWorkbook(excelFile);
            }
        } catch (Exception err) {
            err.printStackTrace();
        }
        try {

//POIFSFileSystem poifs = new POIFSFileSystem(excelFile) ;
//          HSSFWorkbook hworkbook = new HSSFWorkbook(poifs);
            int numberofsheets = workbook.getNumberOfSheets();
            //int numberofsheets = 13;
            for (int sn =0; sn < numberofsheets; sn++) {

                XSSFSheet hsheet = (XSSFSheet) workbook.getSheetAt(sn);
//                int idColumn1 = 14;
//                int idColumn2 = 16;
//                int pictureColumn = 0;
                //HSSFSheet sheet = null;
                int pcount = 0;
                sheetname = hsheet.getSheetName();

                for (XSSFShape shape : hsheet.getDrawingPatriarch().getShapes()) {
                    if (shape instanceof XSSFPicture) {
                        XSSFPicture picture = (XSSFPicture) shape;
                        XSSFClientAnchor anchor = (XSSFClientAnchor) picture.getAnchor();

                        // Ensure to use only relevant pictures
//                  if (anchor.getCol1() == pictureColumn) {
                        // Use the row from the anchor
                        int pCol = anchor.getCol1();
                        XSSFRow pRow = hsheet.getRow(anchor.getRow1());
                        if (pRow != null) {
                            /*                          XSSFCell idCell14 = pictureRow.getCell(17);
                          XSSFCell idCell16 = pictureRow.getCell(19);
                         System.out.println(idCell14);*/
                            int row = pRow.getRowNum();//(int)idCell14.getNumericCellValue();
                            int col = pCol;//(int)idCell16.getNumericCellValue();

                            System.out.println(row + ":" + col);
                            XSSFPictureData data = picture.getPictureData();
                            byte data1[] = data.getData();
                            int pictype = data.getPictureType();
                            iRC[row][col] = 1;
                            String picName = PIC_FILE_LOCATION + XML_ROOT + "-" + sheetname + "-" + row + "." + col + ".png";
                            FileOutputStream out = new FileOutputStream(picName);
                            out.write(data1);
                            out.close();

                            PIC sPIC = new PIC(row, col, picName);
                            listPIC.add(sPIC);

                            pcount++;
                        }
//                  }
                    }
                }
//            }
//            } catch (Exception e) {
//            String er = e.toString();
//        }
//        try{

                /*            
List lst = workbook.getAllPictures();
 for (Iterator it = lst.iterator(); it.hasNext(); ) {
   PictureData pict = (PictureData)it.next();
   String ext = pict.suggestFileExtension();
   byte[] data = pict.getData();
   if (ext.equals("jpeg") || ext.equals("wdp")){
    FileOutputStream out = new FileOutputStream("pict.jpg");
    out.write(data);
    out.close();
   }
   else if (ext.equals("png")){
    FileOutputStream out = new FileOutputStream("pict.png");
    out.write(data);
    out.close();
   }

 }            
                 */
                org.apache.poi.ss.usermodel.Sheet datatypeSheet = workbook.getSheetAt(sn);
                Iterator<Row> iterator = datatypeSheet.iterator();
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder;

            dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.newDocument();
            //add elements to Document
            Element rootElement = doc.createElementNS(NS_URL, XML_ROOT);

                //append root element to document
                doc.appendChild(rootElement);
                Cell currentCell;
                String cellS[] = {"", "", "", "", "", "", "", "", "", "", "", ""};
                int iC = 0;
                iterator.hasNext();
                Row currentRow = iterator.next();
//            currentRow = iterator.next();
//            currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext() && iC < 12) {
                    currentCell = cellIterator.next();
                    if (currentCell.getCellType() == CellType.STRING) {
                        cell0S[iC] = currentCell.getStringCellValue().replaceAll("\\s", "");
                        System.out.print(cell0S[iC] + "--");
                    } else if (currentCell.getCellType() == CellType.NUMERIC) {
                        cell0S[iC] = String.valueOf(currentCell.getNumericCellValue());
                        System.out.print(cell0S[iC] + "--");
                    }
                    iC++;
                }
                //cell0S[1] += "-";
                System.out.println("\r\nEnd of Title Row\r\n");

                int rownum = 0;
                int cN = 0;
                while (iterator.hasNext() && cN != 99999) {

                    currentRow = iterator.next();
                    rownum = currentRow.getRowNum();
                    //cellIterator = currentRow.iterator();
                    cellIterator = currentRow.cellIterator();//.iterator();
                    iC = 0;
                    cN = 0;
                    cellS[iC] = "";
                    while (cellIterator.hasNext() && iC < 12) {

                        currentCell = cellIterator.next();
                        if (currentCell.getCellType() == CellType.STRING) {
                            cellS[iC] = currentCell.getStringCellValue();
                            System.out.print(cellS[iC] + "S--\r\n");
                        } else if (currentCell.getCellType() == CellType.NUMERIC) {
                            if (iC == 0) {
                                cN = (int) currentCell.getNumericCellValue();
                                if (cN == 0) {
                                    break;
                                }
                            }
                            cellS[iC] = String.valueOf(currentCell.getNumericCellValue());
                            System.out.print(currentCell.getNumericCellValue() + "N--\r\n");
                        } else if (currentCell.getCellType() == CellType.BLANK) {
                            String imgPath = "<!--@#@-->";// "<!--IMAGE-->";
                            //<!--@#@--><img src="../oodesign/image001.gif"></img>
                            // int rownum = currentRow.getRowNum();
                            System.out.println(rownum);
                            int cellnum = currentCell.getColumnIndex();
                            if (iRC[rownum][cellnum] == 1) {
                                //cellS[iC] = String.valueOf(imgPath + "<video controls> <source src =\"../picturesforQB/" + XML_ROOT + "-" + sheetname + "-" + rownum + "." + cellnum + ".png\"" + " " + "type=\"audio/x-wav\"></video>");
                                cellS[iC] = String.valueOf(imgPath + "<img src=\"../picturesforQB/" + XML_ROOT + "-" + sheetname + "-" + rownum + "." + cellnum + ".png\""+"></img>");
                            } else {
                                cellS[iC] = " ";
                            }

                        }

                        iC++;
                    }

                    if (cN > 0) {
                        rootElement.appendChild(getQuestion(doc, String.valueOf(cN), cellS[1], cellS[2], cellS[3], cellS[4], cellS[5], cellS[6], cellS[7], cellS[8], cellS[9], cellS[10], cellS[11]));
                    }

                    System.out.println();
                }

                  //for output to file, console
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            //for pretty print
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            DOMSource source = new DOMSource(doc);

            //write to console or file
            StreamResult console = new StreamResult(System.out);
            //StreamResult file = new StreamResult(new File(XML_FILE_LOCATION));
            StreamResult file = new StreamResult(new File("C:\\Users\\giri\\Downloads\\"+sheetname+".xml"));
            

            //write data
            transformer.transform(source, console);
            transformer.transform(source, file);
            System.out.println("DONE");
            }

           /* //for output to file, console
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            //for pretty print
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            DOMSource source = new DOMSource(doc);

            //write to console or file
            StreamResult console = new StreamResult(System.out);
            //StreamResult file = new StreamResult(new File(XML_FILE_LOCATION));
            StreamResult file = new StreamResult(new File(XML_FILE_LOCATION));

            //write data
            transformer.transform(source, console);
            transformer.transform(source, file);
            System.out.println("DONE");*/

        } catch (Exception e) {
            String er = e.toString();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }
    }

    private static Node getQuestion(Document doc, String id, String Val1, String Val2, String Val3,
            String Val4, String Val5, String Val6, String Val7, String Val8, String Val9, String Val10, String Val11) {
        Element question = doc.createElement(SubjectName);
        //String cell0Ss = cell0S[1]+id;

        //set id attribute
        // question.setAttribute(cell0S[0], id);
        //create name element
        //question.appendChild(getQuestionElements(doc, question, cell0S[0], "Question "+id));
        if (Val1 == "") {
            Val1 = " ";
        }
        if (Val2 == "") {
            Val2 = " ";
        }
        if (Val3 == "") {
            Val3 = " ";
        }
        if (Val4 == "") {
            Val4 = " ";
        }
        if (Val5 == "") {
            Val5 = " ";
        }
        if (Val6 == "") {
            Val6 = " ";
        }
        if (Val7 == "") {
            Val7 = " ";
        }
        if (Val8 == "") {
            Val8 = " ";
        }
        if (Val9 == "") {
            Val9 = " ";
        }
        if (Val10 == "") {
            Val10 = " ";
        }
        if (Val11 == "") {
            Val11 = " ";
        }
        question.appendChild(getQuestionElements(doc, question, cell0S[1], Val1));
        question.appendChild(getQuestionElements(doc, question, cell0S[2], Val2));
        question.appendChild(getQuestionElements(doc, question, cell0S[3], Val3));
        question.appendChild(getQuestionElements(doc, question, cell0S[4], Val4));
        question.appendChild(getQuestionElements(doc, question, cell0S[5], Val5));
        question.appendChild(getQuestionElements(doc, question, cell0S[6], Val6));
        question.appendChild(getQuestionElements(doc, question, cell0S[7], Val7));
        question.appendChild(getQuestionElements(doc, question, cell0S[8], Val8));
        question.appendChild(getQuestionElements(doc, question, cell0S[9], Val9));
        question.appendChild(getQuestionElements(doc, question, cell0S[10], Val10));
        question.appendChild(getQuestionElements(doc, question, cell0S[11], Val11));

        return question;
    }

    //utility method to create text node
    private static Node getQuestionElements(Document doc, Element element, String name, String value) {
//            System.out.println("DONE::::"+name);
        Element node = doc.createElement(name);
        node.appendChild(doc.createTextNode(value));
        return node;
    }

}
/*            
                        workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));

            Sheet sheet = workbook.getSheet(0);

            int rows=0;
            int cols=0;
            rows=sheet.getRows();
            cols=sheet.getColumns();
            System.out.print("rows" + ":" + rows + "cols" + ":" + cols + "\r\n");    // Rows, columns
            
            int iR=0;
            cell00S = sheet.getCell(0, iR).getContents();
            cell10S = sheet.getCell(1, iR).getContents();
            cell20S = sheet.getCell(2, iR).getContents();
            cell30S = sheet.getCell(3, iR).getContents();
            cell40S = sheet.getCell(4, iR).getContents();
            cell50S = sheet.getCell(5, iR).getContents();
            cell60S = sheet.getCell(6, iR).getContents();
            cell70S = sheet.getCell(7, iR).getContents();
            cell80S = sheet.getCell(8, iR).getContents();                
                iR = 1;
            for (int iC=0; iR < rows; iR++) {
                String cell0S = sheet.getCell(0, iR).getContents();
                if ( cell0S != null && cell0S.length() != 0) {
                    String cell1S = sheet.getCell(1, iR).getContents();
                    String cell2S = sheet.getCell(2, iR).getContents();
                    String cell3S = sheet.getCell(3, iR).getContents();
                    String cell4S = sheet.getCell(4, iR).getContents();
                    String cell5S = sheet.getCell(5, iR).getContents();
                    String cell6S = sheet.getCell(6, iR).getContents();
                    String cell7S = sheet.getCell(7, iR).getContents();
                    String cell8S = sheet.getCell(8, iR).getContents();
                    System.out.print(cell00S + ":" + cell10S + ":" + cell20S + ":" + cell30S + ":" + cell40S +
                            ":" + cell50S + ":" + cell60S + ":" + cell70S + ":" + cell80S + "\r\n");    // Data
                    System.out.print(cell0S + ":" + cell1S + ":" + cell2S + ":" + cell3S + ":" + cell4S +
                            ":" + cell5S + ":" + cell6S + ":" + cell7S + ":" + cell8S + "\r\n");    // Data

                    //append second child
//                    rootElement.appendChild(getQuestion(doc, "2", "Lisa", "35", "Manager", "Female"));
                    rootElement.appendChild(getQuestion(doc, cell0S, cell1S, cell2S, cell3S, cell4S, cell5S, cell6S));

            Cell cell1 = sheet.getCell(0, 0);
            System.out.print(cell1.getContents() + ":");    // Test Count + :
            Cell cell2 = sheet.getCell(0, 1);
            System.out.println(cell2.getContents());        // 1

            Cell cell3 = sheet.getCell(1, 0);
            System.out.print(cell3.getContents() + ":");    // Result + :
            Cell cell4 = sheet.getCell(1, 1);
            System.out.println(cell4.getContents());        // Passed

            System.out.print(cell1.getContents() + ":");    // Test Count + :
            cell2 = sheet.getCell(0, 2);
            System.out.println(cell2.getContents());        // 2

            System.out.print(cell3.getContents() + ":");    // Result + :
            cell4 = sheet.getCell(1, 2);
            System.out.println(cell4.getContents());        // Passed 2
 */
