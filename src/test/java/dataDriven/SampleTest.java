package dataDriven;

import java.io.IOException;
import java.util.ArrayList;

public class SampleTest {

	public static void main(String[] args) throws IOException {
		
		DataDrivenTest ddt = new DataDrivenTest();
		ArrayList<String> data = ddt.getDataFromExcel("purchase");
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		System.out.println(data.get(4));

	}

}
