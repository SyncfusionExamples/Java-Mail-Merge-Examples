package mailmerge;

import java.io.File;
import com.syncfusion.docio.*;

public class MailMergeWordDocument {

	public static void main(String[] args) throws Exception {
		// Creates a WordDocument instance
		WordDocument document = new WordDocument();
		String basePath = getDataDir("LetterTemplate.docx");
		// Opens the template Word document.
		document.open(basePath, FormatType.Docx);
		String[] fieldNames = { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
		String[] fieldValues = { "Nancy Davolio", "Syncfusion", "507 - 20th Ave. E.Apt. 2A", "Seattle, WA", "USA",
				"(206) 555-9857-x5467" };
		// Performs mail merge
		document.getMailMerge().execute(fieldNames, fieldValues);
		// Saves the Word document
		document.save("Result.docx");
		// Closes the Word document
		document.close();

		System.out.println("Document generated successfully");
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
