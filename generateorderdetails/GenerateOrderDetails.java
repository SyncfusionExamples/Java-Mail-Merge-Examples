package generateorderdetails;

import java.io.FileInputStream;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.RefSupport;
import com.syncfusion.javahelper.system.StringSupport;
import com.syncfusion.javahelper.system.collections.DictionaryEntrySupport;
import com.syncfusion.javahelper.system.collections.generic.DictionarySupport;
import com.syncfusion.javahelper.system.collections.generic.IDictionarySupport;
import com.syncfusion.javahelper.system.collections.generic.IEnumeratorSupport;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.XmlDocumentSupport;
import com.syncfusion.javahelper.system.xml.XmlNodeSupport;
import com.syncfusion.javahelper.system.xml.XmlNodeType;
import com.syncfusion.javahelper.system.xml.XmlReaderSupport;
import java.io.File;

public class GenerateOrderDetails 
{

	public static void main(String[] args) throws Exception 
	{
		//Creates new Word document instance for Word processing.
		WordDocument document = new WordDocument();
		//Opens the template Word document.
		String basePath = getDataDir("Template.docx");
		document.open(basePath, FormatType.Docx);
		//Retrieves the mail merge data.
		MailMergeDataTable dataTable = getMailMergeDataTable();
		//Executes nested Mail merge using implicit relational data.
		document.getMailMerge().executeNestedGroup(dataTable);
		//Removes empty page at the end of Word document.
		removeEmptyPage(document);
		//Save the document in the given name and format.
		document.save("Sample.docx", FormatType.Docx);
		//Release the resources occupied by the WordDocument instance.
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 * @return
	 */
	private static MailMergeDataTable getMailMergeDataTable() throws Exception 
	{
		//Gets the customer details implicit as "IEnumerable" collection.
		ListSupport<CustomerDetails> customers = new ListSupport<CustomerDetails>(CustomerDetails.class);

		FileStreamSupport stream = new FileStreamSupport(getDataDir("CustomerDetails.xml"), FileMode.OpenOrCreate);

		XmlReaderSupport reader = XmlReaderSupport.create(stream);

		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();

		reader.read();

		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();

		while (!(reader.getLocalName() == "CustomerDetails")) 
		{
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) 
			{
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) 
				{
					case "Customers":
						customers.add(getCustomer(reader));
						break;
				}
			} 
			else
			{
				reader.read();
				if ((reader.getLocalName() == "CustomerDetails")
					&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}

		reader.close();
		stream.close();

		// Creates an instance of "MailMergeDataTable" by specifying mail merge group
		// name and "IEnumerable" collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Customers", customers);
		return dataTable;
	}

	/**
	 * 
	 * Gets the customer.
	 * 
	 * @param reader The reader.
	 * @return
	 */
	private static CustomerDetails getCustomer(XmlReaderSupport reader) throws Exception 
	{
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();

		reader.read();

		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();

		CustomerDetails customer = new CustomerDetails();

		while (!(reader.getLocalName() == "Customers")) 
		{
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) 
			{
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) 
				{
					case "CustomerName":
						customer.setCustomerName(reader.readContentAsString());
						break;
					case "Address":
						customer.setAddress(reader.readContentAsString());
						break;
					case "City":
						customer.setCity(reader.readContentAsString());
						break;
					case "PostalCode":
						customer.setPostalCode(reader.readContentAsString());
						break;
					case "Country":
						customer.setCountry(reader.readContentAsString());
						break;
					case "Phone":
						customer.setPhone(reader.readContentAsString());
						break;
					case "Orders":
						customer.getOrders().add(getOrder(reader));
						break;
					default:
						reader.skip();
						break;
				}
			} 
			else
			{
				reader.read();
				if ((reader.getLocalName() == "Customers")
					&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return customer;
	}

	/**
	 * 
	 * Gets the order.
	 * 
	 * @param reader The reader.
	 * @return
	 */
	private static OrderDetails getOrder(XmlReaderSupport reader) throws Exception 
	{
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();

		reader.read();

		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();

		OrderDetails order = new OrderDetails();

		while (!(reader.getLocalName() == "Orders")) 
		{
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) 
			{
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) 
				{
				    case "CustomerName":
						order.setCustomerName(reader.readContentAsString());
						break;
					case "OrderID":
						order.setOrderID(reader.readContentAsString());
						break;
					case "OrderDate":
						order.setOrderDate(reader.readContentAsString());
						break;
					case "ExpectedDeliveryDate":
						order.setExpectedDeliveryDate(reader.readContentAsString());
						break;
					case "ShippedDate":
						order.setShippedDate(reader.readContentAsString());
						break;
					case "Products":
						order.getProducts().add(getProduct(reader));
						break;
				}
				reader.read();
			} 
			else
			{
				reader.read();
				if ((reader.getLocalName() == "Orders")
					&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return order;
	}

	/**
	 * 
	 * Gets the product.
	 * 
	 * @param reader The reader.
	 * @return
	 */
	private static ProductDetails getProduct(XmlReaderSupport reader) throws Exception 
	{
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();

		reader.read();

		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();

		ProductDetails product = new ProductDetails();

		while (!(reader.getLocalName() == "Products")) 
		{
			if (reader.getNodeType().getEnumValue() != XmlNodeType.EndElement.getEnumValue()) 
			{
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) 
				{
					case "OrderID":
						product.setOrderID(reader.readContentAsString());
						break;
					case "Product":
						product.setProduct(reader.readContentAsString());
						break;
					case "UnitPrice":
						product.setUnitPrice(reader.readContentAsString());
						break;
					case "Quantity":
						product.setQuantity(reader.readContentAsString());
						break;
				}
				reader.read();
			} 
			else
			{
				reader.read();
				if ((reader.getLocalName() == "Products")
					&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return product;
	}

	/**
	 * 
	 * Removes empty paragraphs from the end of Word document.
	 * 
	 * @param document The Word document
	 */
	private static void removeEmptyPage(WordDocument document) throws Exception 
	{
		WTextBody textBody = document.getLastSection().getBody();

		//A flag to determine any renderable item found in the Word document.
		boolean IsRenderableItem = false;
		//Iterates text body items.
		for (int itemIndex = textBody.getChildEntities().getCount() - 1; itemIndex >= 0 && !IsRenderableItem; itemIndex--) 
		{
		 	//Check item is empty paragraph and removes it.
			if (textBody.getChildEntities().get(itemIndex) instanceof WParagraph)
			{
				WParagraph paragraph = (WParagraph) textBody.getChildEntities().get(itemIndex);
				//Iterates into paragraph
				for (int pIndex = paragraph.getItems().getCount() - 1; pIndex >= 0; pIndex--) 
				{
					ParagraphItem paragraphItem = paragraph.getItems().get(pIndex);

					//If page break found in end of document, then remove it to preserve contents in same page
					if ((paragraphItem instanceof Break 
					&& ((Break) paragraphItem).getBreakType().getEnumValue() == BreakType.PageBreak.getEnumValue()))
					{
						paragraph.getItems().removeAt(pIndex);
					}
					//Check paragraph contains any renderable items.
					else if (!(paragraphItem instanceof BookmarkStart || paragraphItem instanceof BookmarkEnd)) 
					{
						IsRenderableItem = true;
						//Found renderable item and break the iteration.
						break;
					}
				}
				//Remove empty paragraph and the paragraph with bookmarks only
				if (paragraph.getItems().getCount() == 0 || !IsRenderableItem)
					textBody.getChildEntities().removeAt(itemIndex);
			}
		}
	}

	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	public static String getDataDir(String path) 
	{
		File dir = new File(System.getProperty("user.dir"));
		if (!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}

}
