import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "address", propOrder = {
    "address"
})
public class Server {
    
    @XmlElement(name = "address", required = true)
    protected String address;
    
    @Override
    public String toString() {
        return address;
    }
    
    public void setAddress(String address){
    	this.address = address;
    }
}
