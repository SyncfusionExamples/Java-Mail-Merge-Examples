package generateorderdetails;

import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class OrderDetails {
	private String m_customerName;
	private String m_orderID;
	private String m_orderDate;
	private String m_expectedDeliveryDate;
	private String m_shippedDate;
	private ListSupport<ProductDetails> m_products;

	public String getCustomerName() throws Exception {
		return m_customerName;
	}

	public String setCustomerName(String value) throws Exception {
		m_customerName = value;
		return value;
	}

	public String getOrderID() throws Exception {
		return m_orderID;
	}

	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	public String getOrderDate() throws Exception {
		return m_orderDate;
	}

	public String setOrderDate(String value) throws Exception {
		m_orderDate = value;
		return value;
	}

	public String getExpectedDeliveryDate() throws Exception {
		return m_expectedDeliveryDate;
	}

	public String setExpectedDeliveryDate(String value) throws Exception {
		m_expectedDeliveryDate = value;
		return value;
	}

	public String getShippedDate() throws Exception {
		return m_shippedDate;
	}

	public String setShippedDate(String value) throws Exception {
		m_shippedDate = value;
		return value;
	}

	public ListSupport<ProductDetails> getProducts() throws Exception {
		if (m_products == null)
			m_products = new ListSupport<ProductDetails>(ProductDetails.class);
		return m_products;
	}

	public ListSupport<ProductDetails> setProducts(ListSupport<ProductDetails> value) throws Exception {
		m_products = value;
		return value;
	}

	public OrderDetails() throws Exception {
		m_products = new ListSupport<ProductDetails>(ProductDetails.class);
	}
}
