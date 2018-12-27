import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "servers")
public class Servers {
    
    @XmlElement(name = "server")
    protected List<Server> servers;
    
    public List<Server> getServers() {
        if (servers == null) {
            servers = new ArrayList<>();
        }
        
        return this.servers;
    }
}
