package trial;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

@Test
public class test {

	public void test1() throws IOException {

		String path = "";

		xlutility xl = new xlutility(path);

		int rowcount = xl.gettotlCountRow("Sheet1");

		int coulmcount = xl.gettotalCountColum("Sheet1");

		System.out.println(rowcount);
		System.out.println(coulmcount);

	}

}
