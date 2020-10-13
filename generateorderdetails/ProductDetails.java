package generateorderdetails;

public class ProductDetails {
	private String m_orderID;
	private String m_product;
	private String m_unitPrice;
	private String m_quantity;

	public String getOrderID() throws Exception {
		return m_orderID;
	}

	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	public String getProduct() throws Exception {
		return m_product;
	}

	public String setProduct(String value) throws Exception {
		m_product = value;
		return value;
	}

	public String getUnitPrice() throws Exception {
		return m_unitPrice;
	}

	public String setUnitPrice(String value) throws Exception {
		m_unitPrice = value;
		return value;
	}

	public String getQuantity() throws Exception {
		return m_quantity;
	}

	public String setQuantity(String value) throws Exception {
		m_quantity = value;
		return value;
	}
}
