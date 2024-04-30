package docx4j.testing;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;


@SpringBootApplication
public class TestingApplication {

	public static void main(String[] args) throws Exception {
		ConfigurableApplicationContext context = SpringApplication.run(TestingApplication.class, args);
		DocumentService documentService = context.getBean(DocumentService.class);

		User user = new User();
		user.setFirstName("John");
		user.setLastName("Doe");
		user.setSalutation("This is John Doe");
		user.setMessage("This is Message John Doe");

		documentService.generateDocxFileFromTemplate(user);
	}

}
