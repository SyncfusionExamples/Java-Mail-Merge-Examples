package generateorderdetails;

import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class CustomerDetails {
	private String m_customerName;
	private String m_address;
	private String m_city;
	private String m_postalCode;
	private String m_country;
	private String m_phone;
	private ListSupport<OrderDetails> m_orders;

	public String getCustomerName() throws Exception {
		return m_customerName;
	}

	public String setCustomerName(String value) throws Exception {
		m_customerName = value;
		return value;
	}

	public String getAddress() throws Exception {
		return m_address;
	}

	public String setAddress(String value) throws Exception {
		m_address = value;
		return value;
	}

	public String getCity() throws Exception {
		return m_city;
	}

	public String setCity(String value) throws Exception {
		m_city = value;
		return value;
	}

	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
		return value;
	}

	public String getCountry() throws Exception {
		return m_country;
	}

	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}

	public String getPhone() throws Exception {
		return m_phone;
	}

	public String setPhone(String value) throws Exception {
		m_phone = value;
		return value;
	}

	public ListSupport<OrderDetails> getOrders() throws Exception {
		if (m_orders == null)
			m_orders = new ListSupport<OrderDetails>(OrderDetails.class);
		return m_orders;
	}

	public ListSupport<OrderDetails> setOrders(ListSupport<OrderDetails> value) throws Exception {
		m_orders = value;
		return value;
	}

	public CustomerDetails() throws Exception {
		m_orders = new ListSupport<OrderDetails>(OrderDetails.class);
	}
}
