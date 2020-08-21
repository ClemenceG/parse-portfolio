package com.personal.parse_benchmark_clean;

import java.util.Date;

public class PriceElement {
	private Date date;
	private double price;
	//private int aggVolume;
	
	PriceElement(Date date, double price) {
		this.date = date;
		this.price = price;
	}
	
	public Date getDate() {
		return this.date;
	}
	public double getPrice() {
		return this.price;
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((date == null) ? 0 : date.hashCode());
		long temp;
		temp = Double.doubleToLongBits(price);
		result = prime * result + (int) (temp ^ (temp >>> 32));
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
		PriceElement other = (PriceElement) obj;
		if (date == null) {
			if (other.date != null)
				return false;
		} else if (!date.equals(other.date))
			return false;
		if (Double.doubleToLongBits(price) != Double.doubleToLongBits(other.price))
			return false;
		return true;
	}
	@Override
	public String toString() {
		return "DailyInfo [date=" + date + ", price=" + price + "]";
	}
	
	
}
