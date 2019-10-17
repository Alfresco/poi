package org.apache.poi.hpsf;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

import junit.framework.TestCase;

public class TestVector extends TestCase {

	public void test61295() throws Exception {
		boolean passed = false;
		File initialFile = new File("src/test/resources/61295.tmp");
		InputStream stream = new FileInputStream(initialFile);
		NPOIFSFileSystem npoifs = new NPOIFSFileSystem(stream);
		final DirectoryNode root = npoifs.getRoot();
		DocumentEntry entry = (DocumentEntry) root.getEntry("\005DocumentSummaryInformation");
		
		try {
			PropertySet properties = new PropertySet(new DocumentInputStream(entry));
		} catch (ArrayIndexOutOfBoundsException e) {
			passed = true;
		}
		npoifs.close();
		assertTrue("No ArrayIndexOutOfBoundsException has been thrown", passed);
	}
}