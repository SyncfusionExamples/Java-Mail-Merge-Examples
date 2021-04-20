import java.io.File;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class CreateEnvelopes {

	public static void main(String[] args) throws Exception {
		// Opens the template document.
		WordDocument document = new WordDocument(getDataDir("Envelope_Template.docx"), FormatType.Docx);
		// Gets the organization details as “IEnumerable” collection.
		ListSupport<Recipient> recipients = getRecipient();
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group
		// name and “IEnumerable” collection.
		document.getMailMerge().execute(recipients);
		//Saves and closes the document instance
		document.save("Result.docx", FormatType.Docx);
		document.close();
	}

	public static ListSupport<Recipient> getRecipient() throws Exception {
		// Creates Employee details.
		ListSupport<Recipient> recipients = new ListSupport<Recipient>();
		recipients.add(new Recipient("Nancy", "Davolio", "507 - 20th Ave. E.Apt. 2A", "Seattle", "WA", "98122"));
		recipients.add(new Recipient("Andrew", "Fuller", "908 W. Capital Way", "Tacoma", "WA", "98401"));
		recipients.add(new Recipient("Janet", "Leverling", "722 Moss Bay Blvd.", "Kirkland", "WA", "98033"));
		recipients.add(new Recipient("Margaret", "Peacock", "4110 Old Redmond Rd.", "Redmond", "WA", "98052"));
		recipients.add(new Recipient("Steven", "Buchanan", "14 Garrett Hil", "London", "", "SW1 8JR"));
		return recipients;
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
