package com.personal.parse_benchmark_clean;

public class Sector {
	private int sectorNb;
	private String sectorName;
	
	Sector(int sectorNb) {
		this.sectorNb = sectorNb;
	}
	Sector(String sectorName) {
		this.sectorName = sectorName;
	}
	
	public int getSectorNb() {
		return this.sectorNb;
	}
	public String getSectorName() {
		return this.sectorName;
	}
	
	public void addSectorNb(int sectorNb) {
		this.sectorNb = sectorNb;
	}
	
	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((sectorName == null) ? 0 : sectorName.hashCode());
		result = prime * result + sectorNb;
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
		Sector other = (Sector) obj;
		if (sectorName == null) {
			if (other.sectorName != null)
				return false;
		} else if (!sectorName.equals(other.sectorName))
			return false;
		if (sectorNb != other.sectorNb)
			return false;
		return true;
	}
	@Override
	public String toString() {
		return "Sector [sectorNb=" + sectorNb + ", sectorName=" + sectorName + "]";
	}
	

}
