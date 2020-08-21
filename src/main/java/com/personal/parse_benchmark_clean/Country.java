package com.personal.parse_benchmark_clean;

public class Country {
	private String CountryName;
	private String CountryIso;
	
	Country(String countryName) {
		this.CountryName = countryName;
	}
	
	public String getCountryIso() {
		return this.CountryIso;
	}
	public String getCountryName() {
		return this.CountryName;
	}
	
	public void setCountryIso(String countryIso) {
		this.CountryIso = countryIso;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((CountryIso == null) ? 0 : CountryIso.hashCode());
		result = prime * result + ((CountryName == null) ? 0 : CountryName.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Country other = (Country) obj;
		if (CountryIso == null) {
			if (other.CountryIso != null)
				return false;
		} else if (!CountryIso.equals(other.CountryIso))
			return false;
		if (CountryName == null) {
			if (other.CountryName != null)
				return false;
		} else if (!CountryName.equals(other.CountryName))
			return false;
		return true;
	}

	@Override
	public String toString() {
		return "Country [CountryName=" + CountryName + ", CountryIso=" + CountryIso + "]";
	}
	
	
	
}
