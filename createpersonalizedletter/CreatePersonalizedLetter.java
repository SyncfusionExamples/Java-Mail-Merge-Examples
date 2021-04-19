import java.io.*;
import com.syncfusion.docio.*;

public class CreatePersonalizedLetter {

	public static void main(String[] args) throws Exception {
		// Opens the template document.
		WordDocument document = new WordDocument(getDataDir("LetterTemplate.docx"), FormatType.Docx);
		String[] fieldNames = new String[] { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
		String[] fieldValues = new String[] { "Nancy Davolio", "Syncfusion", "507 - 20th Ave. E.Apt. 2A", "Seattle, WA",
				"USA", "(206) 555-9857-x5467" };
		// Performs the mail merge.
		document.getMailMerge().execute(fieldNames, fieldValues);
		// Saves the Word document.
		document.save("Result.docx", FormatType.Docx);
		// Closes the Word document.
		document.close();
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
