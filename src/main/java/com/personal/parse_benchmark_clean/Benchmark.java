package com.personal.parse_benchmark_clean;

import java.util.List;
import java.util.ArrayList;

public class Benchmark {
	private String name;
	private List<BenchmarkElement> benchmarkElements = new ArrayList<BenchmarkElement>();
	
	Benchmark(String name) {
		this.name = name;
	}
	
	
	public String getSecurityTicker(int i) {
		return this.benchmarkElements.get(i).getSecurity().getTicker();
	}
	public String getSecurityBBGlobal(int i) {
		return this.benchmarkElements.get(i).getSecurity().getBloombergGlobal();
	}
	public List<BenchmarkElement> getBenchmarkElements() {
		return this.benchmarkElements;
	}
	public BenchmarkElement getBenchmarkElement(int i) {
		return this.benchmarkElements.get(i);
	}
	public Security getSecurity(int index) {
		return this.benchmarkElements.get(index).getSecurity();
	}
	
	
	public void addBenchmarkElement(BenchmarkElement benchmarkElement) {
		this.benchmarkElements.add(benchmarkElement);
	}
	public void removeBenchmarkElement(int index) {
		this.benchmarkElements.remove(index);
	}
	public void addBenchmarkElement(int index, BenchmarkElement benchmarkElement) {
		this.benchmarkElements.add(index, benchmarkElement);
	}
	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((benchmarkElements == null) ? 0 : benchmarkElements.hashCode());
		result = prime * result + ((name == null) ? 0 : name.hashCode());
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
		Benchmark other = (Benchmark) obj;
		if (benchmarkElements == null) {
			if (other.benchmarkElements != null)
				return false;
		} else if (!benchmarkElements.equals(other.benchmarkElements))
			return false;
		if (name == null) {
			if (other.name != null)
				return false;
		} else if (!name.equals(other.name))
			return false;
		return true;
	}
	@Override
	public String toString() {
		return "Benchmark [name=" + name + ", benchmarkElements=" + benchmarkElements + "]";
	}
	
	
	
}
