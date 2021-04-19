import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import javax.mail.*;
import javax.mail.internet.*;
import org.json.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.*;

public class CreateAndSendMail {

	public static void main(String[] args) throws Exception {

		// Opens the template document.
		WordDocument template = new WordDocument(getDataDir("SentMail_Template.docx"), FormatType.Docx);
		ListSupport<Object> recipients = getRecipients();
		for (Object dataRecord_tempObj : recipients) {
			Object dataRecord = dataRecord_tempObj;
			WordDocument document = template.clone();
			document.getMailMerge().execute(new ListSupport<Object>(Object.class) {
				{
					add((DictionarySupport<String, Object>) ObjectSupport.instanceOf(dataRecord,
							DictionarySupport.class));
				}
			});
			ByteArrayOutputStream stream = new ByteArrayOutputStream();
			document.getSaveOptions().setHtmlExportOmitXmlDeclaration(true);
			document.save(stream, FormatType.Html);
			document.close();
			String mailBody = stream.toString();
			if (StringSupport.startsWith(mailBody, "<!DOCTYPE"))
				mailBody = StringSupport.remove(mailBody, 0, 97);
			sendEMail("MailId@live.in", "RecipientMailId@live.in", StringSupport.concat(StringSupport.concat(
					"You order #",
					((DictionarySupport<String, Object>) ObjectSupport.instanceOf(dataRecord, DictionarySupport.class))
							.get("OrderID").toString()),
					" has been shipped"), mailBody);
		}
	}

	/**
	 * Gets the data to perform mail merge.
	 * 
	 * @return
	 * @throws Exception
	 */
	private static ListSupport<Object> getRecipients() throws Exception {
		String path = getDataDir("CustomerDetails.json");
		// Read JSON
		List<String> lines = Files.readAllLines(Paths.get(path), StandardCharsets.UTF_8);
		String jsonString = String.join(System.lineSeparator(), lines);
		// Convert JSON as JSONObject
		JSONObject jsonObject = new JSONObject(jsonString);
		IDictionarySupport<String, Object> data = getData(jsonObject);
		return (ListSupport<Object>) ObjectSupport.instanceOf(data.get("Customers"), ListSupport.class);
	}

	/**
	 * 
	 * Gets data from JSON object.
	 * 
	 * @param jsonObject JSON object.
	 * @return Dictionary of data.
	 */
	private static DictionarySupport<String, Object> getData(JSONObject jsonObject) throws Exception {
		DictionarySupport<String, Object> dictionary = new DictionarySupport<String, Object>(String.class,
				Object.class);

		Iterator<Object> keys = jsonObject.keys();
		while (keys.hasNext()) {

			String key = (String) keys.next();
			Object keyValue = null;

			if (jsonObject.get(key) instanceof JSONArray) {
				JSONArray jsonArray = (JSONArray) jsonObject.get(key);
				keyValue = getData(jsonArray);
			} else if (jsonObject.get(key) instanceof String) {
				keyValue = jsonObject.get(key);
			}
			dictionary.add(key, keyValue);
		}
		return dictionary;
	}

	/**
	 * 
	 * Gets array of items from JSON array.
	 * 
	 * @param jArray JSON array.
	 * @return List of objects.
	 */
	private static ListSupport<Object> getData(JSONArray jsonArray) throws Exception {
		ListSupport<Object> jArrayItems = new ListSupport<Object>();

		for (int i = 0, size = jsonArray.length(); i < size; i++) {
			Object keyValue = null;
			if (jsonArray.get(i) instanceof JSONObject) {

				keyValue = getData((JSONObject) jsonArray.get(i));
				jArrayItems.add(keyValue);
			}
		}
		return jArrayItems;
	}

	/**
	 * Sent the mail to the recipients
	 * 
	 * @param from       Sender of the mail
	 * @param recipients Receiver of the mail
	 * @param subject    Subject of the mail
	 * @param body       Message need to be sent
	 * @throws Exception
	 */
	private static void sendEMail(String from, String recipients, String subject, String body) throws Exception {

		Properties props = new Properties();
		props.put("mail.smtp.host", "smtp.live.com");
		props.put("mail.smtp.port", "587");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.auth", "true");

		Session session = Session.getDefaultInstance(props, new javax.mail.Authenticator() {
			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(from, "password");
			}
		});
		// Compose the message
		MimeMessage message = new MimeMessage(session);
		message.setFrom(new InternetAddress(from));
		message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipients));
		message.setSubject(subject);
		message.setText(body);
		// send the message
		Transport.send(message);
		System.out.println("message sent successfully...");
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
