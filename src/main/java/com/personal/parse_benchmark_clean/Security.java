package com.personal.parse_benchmark_clean;

import java.util.ArrayList;

public class Security {
	private String ticker;
	private String bloombergGlobal;
	private String CodeIsin;
	private ArrayList<PriceElement> priceElements = new ArrayList<PriceElement>();
	private String currency;
	private Sector sector;
	private Sector subsector;
	private Sector supersector;
	private Country country;

	
	Security(String name) {
		this.ticker = name;
	}
	
	Security(String name, String global) {
		this.ticker = name;
		this.bloombergGlobal = global;
	}
	
	public String getTicker() {
		return this.ticker;
	}
	public String getBloombergGlobal() {
		return this.bloombergGlobal;
	}
	public ArrayList<PriceElement> getPrices() {
		return this.priceElements;
	}
	public String getCurrency() {
		return this.currency;
	}
	public Sector getSector() {
		return this.sector;
	}
	public Sector getSubsector() {
		return this.subsector;
	}
	public Sector getSupersector() {
		return this.supersector;
	}
	public Country getCountry() { 
		return this.country;
	}
	public String getIsin() {
		return this.CodeIsin;
	}
	

	public void setCurrency(String curr) {
		this.currency = curr;
	}
	public void addPriceElement(PriceElement daily) {
		this.priceElements.add(daily);
	}
	public void setSector(Sector sector) {
		this.sector = sector;
	}
	public void setSubsector(Sector sector) {
		this.subsector = sector;
	}
	public void setSupersector(Sector sector) {
		this.supersector = sector;
	}
	public void setCountry(Country country) {
		this.country = country;
	}
	public void setIsin(String isin) {
		this.CodeIsin = isin;
	}
	public void setTicker(String ticker) {
		this.ticker = ticker;
	}
	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((CodeIsin == null) ? 0 : CodeIsin.hashCode());
		result = prime * result + ((bloombergGlobal == null) ? 0 : bloombergGlobal.hashCode());
		result = prime * result + ((country == null) ? 0 : country.hashCode());
		result = prime * result + ((currency == null) ? 0 : currency.hashCode());
		result = prime * result + ((priceElements == null) ? 0 : priceElements.hashCode());
		result = prime * result + ((sector == null) ? 0 : sector.hashCode());
		result = prime * result + ((subsector == null) ? 0 : subsector.hashCode());
		result = prime * result + ((supersector == null) ? 0 : supersector.hashCode());
		result = prime * result + ((ticker == null) ? 0 : ticker.hashCode());
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
		Security other = (Security) obj;
		if (CodeIsin == null) {
			if (other.CodeIsin != null)
				return false;
		} else if (!CodeIsin.equals(other.CodeIsin))
			return false;
		if (bloombergGlobal == null) {
			if (other.bloombergGlobal != null)
				return false;
		} else if (!bloombergGlobal.equals(other.bloombergGlobal))
			return false;
		if (country == null) {
			if (other.country != null)
				return false;
		} else if (!country.equals(other.country))
			return false;
		if (currency == null) {
			if (other.currency != null)
				return false;
		} else if (!currency.equals(other.currency))
			return false;
		if (priceElements == null) {
			if (other.priceElements != null)
				return false;
		} else if (!priceElements.equals(other.priceElements))
			return false;
		if (sector == null) {
			if (other.sector != null)
				return false;
		} else if (!sector.equals(other.sector))
			return false;
		if (subsector == null) {
			if (other.subsector != null)
				return false;
		} else if (!subsector.equals(other.subsector))
			return false;
		if (supersector == null) {
			if (other.supersector != null)
				return false;
		} else if (!supersector.equals(other.supersector))
			return false;
		if (ticker == null) {
			if (other.ticker != null)
				return false;
		} else if (!ticker.equals(other.ticker))
			return false;
		return true;
	}

	@Override
	public String toString() {
		return "Security [ticker=" + ticker + ", bloombergGlobal=" + bloombergGlobal + ", CodeIsin=" + CodeIsin
				+ ", country=" + country + ", currency=" + currency + ", sector=" + sector + ", subsector=" + subsector
				+ ", supersector=" + supersector + ", priceElements=" + priceElements + "]";
	}
	
	
	
}
