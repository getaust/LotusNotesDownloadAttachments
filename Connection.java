package de.download.util;

import java.io.ByteArrayInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Enumeration;
import java.util.Vector;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import lotus.domino.Database;
import lotus.domino.NotesFactory;
import lotus.domino.RichTextItem;
import lotus.domino.Session;
import lotus.domino.DocumentCollection;
import lotus.domino.EmbeddedObject;
import lotus.domino.NotesException;

/*
 * Created by T. Aust 2023
 */

public class Connection {

	static String Host = "server:63148";
	static String Database = "IT/ERP/AA.nsf";
	static String User = "userdb";
	static String Password = "pwd";
	static String saveDir = "c:\\TEMP\\";
	static String specificDoc = "Besondere Vertragsbedingungen";
	static boolean findOneDoc = false;
	static String specificAuthor = "Hans";
	static boolean findOneAuthor = false;

	int i = 1;

	public static void main(String[] args) {

		try {
			printf("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
			printf("Open session...");
			Session s = NotesFactory.createSession(Host, User, Password);

			printf(" as User " + User + ": " + s.getCommonUserName());
			Database db = s.getDatabase(null, Database);

			printf("Connecting to Domino Lotus Notes: " + db.getURL());
			printf("Connected!");
			// printf("get " + db.getTitle());
			// printf("get " + db.getDB2Schema());
			// printf("get " + db.getHttpURL());
			// printf("get " + db.getServer());
			// printf("get " + db.getAllDocuments().getCount());

			Connection con = new Connection();
			DocumentCollection dc = db.getAllDocuments();
			lotus.domino.Document doc = dc.getFirstDocument();
			printf("Searching documents...");
			printf("\n++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");

			while (doc != null) {
				String Subject = doc.getItemValueString("Subject");
				if (Subject != null) {
					if (findOneDoc || findOneAuthor) {
						if (findOneDoc && Subject.contains(specificDoc)) {
							con.findDocs(doc, true);
							break;
						}
						if (findOneAuthor && doc.generateXML().contains(specificAuthor))
							con.findDocs(doc, true);
					} else
						con.findDocs(doc, false);
				}
				doc = dc.getNextDocument();
			}

			printf("Finished!");

		} catch (Exception e) {
			printf("TRACE :: main :: " + e.getMessage());
			// e.printStackTrace();
		}

	}

	public void findDocs(lotus.domino.Document doc, boolean shouldPrint) {

		try {
			String Subject = doc.getItemValueString("Subject").trim();
			RichTextItem body = (RichTextItem) doc.getFirstItem("Body");

			Vector<?> v = body.getEmbeddedObjects();
			Enumeration<?> e = v.elements();

			boolean attach = false;

			String AttachmentsName = "";
			while (e.hasMoreElements()) {
				EmbeddedObject eo = (EmbeddedObject) e.nextElement();
				if (eo.getType() == EmbeddedObject.EMBED_ATTACHMENT) {
					if (attach == false)
						printf("Subject: " + Subject);
					attach = true;
					eo.extractFile(saveDir + eo.getSource());

					printf(" file: " + i + " ==> " + eo.getSource());
					AttachmentsName += eo.getSource() + "\n";

					// eo.remove();
					i++;
				}
			}
			if (attach || shouldPrint) {
				printf("URL: " + doc.getHttpURL());
				// saveFile(Subject.replaceAll("[^a-zA-Z0-9]", " ") + "---Body.txt",
				// doc.generateXML());
				// #1 fileName, #2 payload

				String fileNameAdjusted = Subject.replaceAll("[^a-zA-Z0-9]", " ");

				// file: Attachment
				saveFile(fileNameAdjusted + "---Attachments.txt", AttachmentsName);
				// file: Description
				saveFile(fileNameAdjusted + "---Description.txt", body.getText().toString());
				// file: Categories
				saveFile(fileNameAdjusted + "---Categories.txt", getCategories(doc.generateXML()));

				printf("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
			}
		} catch (NotesException | IOException e) {
			printf("TRACE :: findDocs :: " + e.getMessage());
		}

	}

	public void saveFile(String fileName, String payload) throws IOException {
		try {
			if (payload != null && payload.length() > 0) {
				String path = saveDir + fileName;
				Files.write(Paths.get(path), payload.getBytes(StandardCharsets.UTF_8));
				printf(" " + fileName);
			}
		} catch (IOException e) {
			printf("TRACE :: saveFile :: " + e.getMessage());
		}
	}

	public String getCategories(String s) {

		String response = "";
		String rule = "/document/item[starts-with(@name,'GLSTopicsChoices')]";

		try {
			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
			DocumentBuilder db = dbf.newDocumentBuilder();
			Document xml = db.parse(new ByteArrayInputStream(s.getBytes(StandardCharsets.UTF_8)));

			XPath xPath = XPathFactory.newInstance().newXPath();
			XPathExpression expr = xPath.compile(rule);
			NodeList nodes = (NodeList) expr.evaluate(xml, XPathConstants.NODESET);

			for (int i = 0; i < nodes.getLength(); i++) {
				Node node = nodes.item(i);
				String category = node.getTextContent().trim();
				if (category != null && category.length() > 0)
					response += category + "\n";
			}

		} catch (XPathExpressionException | SAXException | IOException | ParserConfigurationException e) {
			printf("TRACE :: getCategories :: " + e.getMessage());
		}
		return response;

	}

	/*
	* show console and save local file
	*/
	public static void printf(String message) {
		try {
			System.out.println(message);
			PrintWriter out = new PrintWriter(new FileWriter(saveDir + "OutputConsole.txt", true), true);
			out.write(message);
			out.write("\r\n");
			out.close();
		} catch (IOException e) {
			e.getStackTrace();
		}
	}

}
