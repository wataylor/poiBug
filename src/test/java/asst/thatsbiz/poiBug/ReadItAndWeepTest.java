package asst.thatsbiz.poiBug;

import static org.junit.Assert.*;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class ReadItAndWeepTest {

  @BeforeClass
  public static void setUpBeforeClass() throws Exception {
  }

  @AfterClass
  public static void tearDownAfterClass() throws Exception {
  }

  @Before
  public void setUp() throws Exception {
  }

  @After
  public void tearDown() throws Exception {
  }

  @Test
  public void testMain() {
    String fn = getClass().getClassLoader().getResource("asst/thatsbiz/poiBug/20191122-RTP-OnlineOrdering-WA-FansRave-Updates.xlsx").getFile();
    System.out.println(fn);
    String[] args = new String[1];
    args[0] = fn;
    ReaditAndWeepMain.main(args);
  }

}