import org.apache.commons.io.IOUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

/**
 * linux或windows命令执行
 */
public class CommandExecute {

    public static void main(String[] args) {
        CommandExecute obj = new CommandExecute();
        String domainName = "www.baidu.com";
        //in mac oxs
        String command = "ping " + domainName;
        //in windows
        //String command = "ping -n 3 " + domainName;
        String output = obj.executeCommand(command);
        System.out.println(output);
    }

    public static String executeCommand(String command) {
        StringBuffer output = new StringBuffer();
        Process p;
        InputStreamReader inputStreamReader = null;
        BufferedReader reader = null;
        try {
            p = Runtime.getRuntime().exec(command);
            p.waitFor();
            inputStreamReader = new InputStreamReader(p.getInputStream(), "UTF-8");
            reader = new BufferedReader(inputStreamReader);
            String line = "";
            while ((line = reader.readLine()) != null) {
                output.append(line + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(reader);
            IOUtils.closeQuietly(inputStreamReader);
        }
        System.out.println(output.toString());
        return output.toString();

    }
}
