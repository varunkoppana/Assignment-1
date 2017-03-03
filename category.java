package varun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import javax.xml.transform.OutputKeys;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;


import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class category {
	
public static void main (String arg[]){
	List<Cell> table = new ArrayList<Cell>();
	
	try {
		
		FileInputStream file = new FileInputStream(new File("C:/Users/varun/Documents/DATA.xls"));
		
		HSSFWorkbook workbook = new HSSFWorkbook(file);

		
		HSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			Iterator<Cell> cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext()) {
				
				Cell cell = cellIterator.next();
				cell.setCellType(Cell.CELL_TYPE_STRING); 
				table.add(cell);
			}
		}
		file.close();	
		
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	

		  try {

		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

		// root elements
		Document doc = docBuilder.newDocument();
		Element rootElement = doc.createElement("categories");
		doc.appendChild(rootElement);
	
	

		// category books
			Element category = doc.createElement("category");
			rootElement.appendChild(category);

			Attr attr = doc.createAttribute("name");
			attr.setValue("Books");
			category.setAttributeNode(attr);
		//category music
			Element category2 = doc.createElement("category");
			rootElement.appendChild(category2);

			Attr attr4 = doc.createAttribute("name");
			attr4.setValue("Music");
			category2.setAttributeNode(attr4);
		
		//category electronics	        
			Element category3 = doc.createElement("category");
			rootElement.appendChild(category3);

			Attr attr8 = doc.createAttribute("name");
			attr8.setValue("Electronics");
			category3.setAttributeNode(attr8);
			
		for(int i=5 ; i < table.size(); i=i+4)
		{
			
			if (table.get(i).toString().equals("Books"))
			{
				Element product = doc.createElement("product");
				category.appendChild(product);
		
				Attr attr2 = doc.createAttribute("id");
				attr2.setValue(table.get(i-1).toString());
				product.setAttributeNode(attr2);
				
				Element name = doc.createElement("name");
				name.appendChild(doc.createTextNode(table.get(i+1).toString()));
				product.appendChild(name);

				Element price = doc.createElement("price");
		        price.appendChild(doc.createTextNode(table.get(i+2).toString()));
		        product.appendChild(price);
			}
			else if(table.get(i).toString().equals("Music"))
			{
				Element product = doc.createElement("product");
				category2.appendChild(product);
		
				Attr attr2 = doc.createAttribute("id");
				attr2.setValue(table.get(i-1).toString());
				product.setAttributeNode(attr2);
				
				Element name = doc.createElement("name");
				name.appendChild(doc.createTextNode(table.get(i+1).toString()));
				product.appendChild(name);

				Element price = doc.createElement("price");
		        price.appendChild(doc.createTextNode(table.get(i+2).toString()));
		        product.appendChild(price);
			}
			else if(table.get(i).toString().equals("Electronics"))
			{
				Element product = doc.createElement("product");
				category3.appendChild(product);
		
				Attr attr2 = doc.createAttribute("id");
				attr2.setValue(table.get(i-1).toString());
				product.setAttributeNode(attr2);
				
				Element name = doc.createElement("name");
				name.appendChild(doc.createTextNode(table.get(i+1).toString()));
				product.appendChild(name);

				Element price = doc.createElement("price");
		        price.appendChild(doc.createTextNode(table.get(i+2).toString()));
		        product.appendChild(price);
			}
		}
		
		// transforming the content into xml file
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "5");
		DOMSource source = new DOMSource(doc);
		StreamResult result = new StreamResult(new File("C:\\XML\\categories.xml"));

		transformer.transform(source, result);

		System.out.println("File saved!");

	}catch (ParserConfigurationException pce) {
		pce.printStackTrace();
	  } catch (TransformerException tfe) {
		tfe.printStackTrace();
	  }
}

}
	
