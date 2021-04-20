public class Recipient {
	private String _firstName;
	private String _lastName;
	private String _address;
	private String _city;
	private String _zipCode;
	private String _country;

	public String getFirstName() throws Exception {
		return _firstName;
	}

	public String setFirstName(String value) throws Exception {
		_firstName = value;
		return value;
	}

	public String getLastName() throws Exception {
		return _lastName;
	}

	public String setLastName(String value) throws Exception {
		_lastName = value;
		return value;
	}

	public String getAddress() throws Exception {
		return _address;
	}

	public String setAddress(String value) throws Exception {
		_address = value;
		return value;
	}

	public String getCity() throws Exception {
		return _city;
	}

	public String setCity(String value) throws Exception {
		_city = value;
		return value;
	}

	public String getZipCode() throws Exception {
		return _zipCode;
	}

	public String setZipCode(String value) throws Exception {
		_zipCode = value;
		return value;
	}

	public String getCountry() throws Exception {
		return _country;
	}

	public String setCountry(String value) throws Exception {
		_country = value;
		return value;
	}

	public Recipient(String firstName, String lastName, String address, String city, String zipcode, String country)
			throws Exception {
		setFirstName(firstName);
		setLastName(lastName);
		setAddress(address);
		setCity(city);
		setZipCode(zipcode);
		setCountry(country);

	}
}
