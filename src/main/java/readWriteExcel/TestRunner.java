package readWriteExcel;

import java.io.IOException;


/**
 * Test runner class to test read write excel functionality
 */
public class TestRunner {

    public static void main(String[] args) throws IOException {
        ReadWriteExcel readWriteExcel=new ReadWriteExcel();
        readWriteExcel.readWriteExcel();
    }
}