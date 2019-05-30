package xyz.wqf31415;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.env.Environment;

import java.net.InetAddress;
import java.net.UnknownHostException;

@SpringBootApplication
public class SpringBootPoiApplication {
	private static final Logger logger = LoggerFactory.getLogger(SpringBootPoiApplication.class);

	public static void main(String[] args) throws UnknownHostException {
		SpringApplication app = new SpringApplication(SpringBootPoiApplication.class);
		Environment env = app.run(args).getEnvironment();
		logger.info("\n==============================================\n" +
				"Application {} is running! Access URL:\n" +
				"\tLocal:\t\thttp://localhost:{}\n" +
				"\tExternal:\thttp://{}:{}\n" +
				"=============================================="
				,env.getProperty("spring.application.name")
				,env.getProperty("server.port")
				,InetAddress.getLocalHost().getHostAddress()
				,env.getProperty("server.port"));
	}
}
