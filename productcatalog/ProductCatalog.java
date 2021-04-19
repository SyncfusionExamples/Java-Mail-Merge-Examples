import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.Int32Support;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.data.*;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.io.*;

public class ProductCatalog{

	public static void main(String[] args) throws Exception {
		// Opens the template document.
		WordDocument document = new WordDocument(getDataDir("MailMergeEvent_Template.docx"), FormatType.Docx);
		// Uses the mail merge events to perform the conditional formatting during
		// runtime.
		DataSetSupport ds = getData();
		document.getMailMerge().MergeField.add("alternateRows_MergeField", new MergeFieldEventHandler() {
			ListSupport<MergeFieldEventHandler> delegateList = new ListSupport<MergeFieldEventHandler>(
					MergeFieldEventHandler.class);

			public void invoke(Object sender, MergeFieldEventArgs args) throws Exception {
				for (MergeFieldEventHandler theDelegate2 : delegateList) {
					AlternateRows_MergeField(sender, args);
				}
				AlternateRows_MergeField(sender, args);
			}

			public void dynamicInvoke(Object... args) throws Exception {
				for (MergeFieldEventHandler theDelegate2 : delegateList) {
					AlternateRows_MergeField((Object) args[0], (MergeFieldEventArgs) args[1]);
				}
				AlternateRows_MergeField((Object) args[0], (MergeFieldEventArgs) args[1]);
			}

			public void add(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}

			public void remove(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
		});
		document.getMailMerge().MergeImageField.add("mergeField_ProductImage", new MergeImageFieldEventHandler() {
			ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
					MergeImageFieldEventHandler.class);

			public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception {
				for (MergeImageFieldEventHandler theDelegate2 : delegateList) {
					mergeField_ProductImage(sender, args);
				}
				mergeField_ProductImage(sender, args);
			}

			public void dynamicInvoke(Object... args) throws Exception {
				for (MergeImageFieldEventHandler theDelegate2 : delegateList) {
					mergeField_ProductImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
				}
				mergeField_ProductImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
			}

			public void add(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}

			public void remove(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
		});
		document.getMailMerge().executeGroup(ds.getTables().get("Products"));
		document.getMailMerge().executeGroup(ds.getTables().get("Product_PriceList"));
		// Saves the Word document.
		document.save("Sample.docx", FormatType.Docx);
		// Closes the Word document.
		document.close();
	}

	/*
	 * Gets the data to perform mail merge
	 */
	private static DataSetSupport getData() throws Exception {
		DataSetSupport ds = new DataSetSupport();
		String[] products = { "Apple Juice", "Grape Juice", "Hot Soup", "Tender Coconut", "Vennila", "Strawberry",
				"Cherry", "Cone" };
		DataRowSupport row;
		DataTableSupport priceListTable = new DataTableSupport("Product_PriceList");
		ds.getTables().add(priceListTable);
		ds.getTables().get(0).getColumns().add("ProductName");
		ds.getTables().get(0).getColumns().add("Price");
		DataTableSupport productsTable = new DataTableSupport("Products");
		ds.getTables().add(productsTable);
		ds.getTables().get(1).getColumns().add("SNO");
		ds.getTables().get(1).getColumns().add("ProductName");
		ds.getTables().get(1).getColumns().add("ProductImage");
		int count = 0;
		for (Object product_tempObj : products) {
			String product = (String) product_tempObj;
			row = ds.getTables().get("Product_PriceList").newRow();
			row.set("ProductName", product);
			switch ((product) == null ? "string_null_value" : (product)) {
			case "Apple Juice":
				row.set("Price", "$12.00");
				break;
			case "Grape Juice":
				row.set("Price", "$15.00");
				break;
			case "Hot Soup":
				row.set("Price", "$20.00");
				break;
			case "Tender coconut":
				row.set("Price", "$10.00");
				break;
			case "Vennila Ice Cream":
				row.set("Price", "$15.00");
				break;
			case "Strawberry":
				row.set("Price", "$18.00");
				break;
			case "Cherry":
				row.set("Price", "$25.00");
				break;
			default:
				row.set("Price", "$20.00");
				break;
			}
			ds.getTables().get("Product_PriceList").getRows().add(row);
			count++;
			row = ds.getTables().get("Products").newRow();
			row.set("SNO", Int32Support.toString(count));
			row.set("ProductName", product);
			row.set("ProductImage", String.valueOf(product).concat(String.valueOf(".png")));
			ds.getTables().get("Products").getRows().add(row);
		}
		return ds;
	}

	/*
	 * Method to handle MergeField event.
	 */
	private static void AlternateRows_MergeField(Object sender, MergeFieldEventArgs args) throws Exception {
		// Sets text color to the alternate mail merge record.
		if (Integer.compare((args.getRowIndex() % 2), 0) == 0) {
			args.getTextRange().getCharacterFormat().setTextColor(ColorSupport.fromArgb(196, 89, 17));
		}
	}

	/*
	 * Method to handle MergeImageField event.
	 */
	private static void mergeField_ProductImage(Object sender, MergeImageFieldEventArgs args) throws Exception {
		// Binds image from file system during mail merge.
		if ((args.getFieldName()).equals("ProductImage")) {
			String ProductFileName = getDataDir(args.getFieldValue().toString());
			// Gets the image from file system.
			FileStreamSupport imageStream = new FileStreamSupport(ProductFileName, FileMode.Open, FileAccess.Read);
			ByteArrayInputStream stream = new ByteArrayInputStream(imageStream.toArray());
			args.setImageStream(stream);
		}
	}

	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	private static String getDataDir(String path) {
		File dir = new File(System.getProperty("user.dir"));
		if (!(dir.toString().endsWith("Java-Mail-Merge-Examples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}
}
