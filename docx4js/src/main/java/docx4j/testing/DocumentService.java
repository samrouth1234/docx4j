package docx4j.testing;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.HashMap;

@Service
public class DocumentService {
	private static final String TEMPLATE_PATH = "template.docx";

	public void generateDocxFileFromTemplate(User user) throws Exception {
		InputStream templateInputStream = this.getClass()
				.getClassLoader()
				.getResourceAsStream(TEMPLATE_PATH);

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);

		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

		VariablePrepare.prepare(wordMLPackage);

		HashMap<String, String> variables = new HashMap<>();
		variables.put("firstName", user.getFirstName());
		variables.put("lastName", user.getLastName());
		variables.put("salutation", user.getSalutation());
		variables.put("message", user.getMessage());

		documentPart.variableReplace(variables);

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		wordMLPackage.save(outputStream);
		outputStream.toByteArray();

		File exportFile = new File("template.docx");
		try {
			wordMLPackage.save(exportFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

