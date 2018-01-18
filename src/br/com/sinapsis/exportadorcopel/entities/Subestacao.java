package br.com.sinapsis.exportadorcopel.entities;

import java.util.HashSet;

public class Subestacao {

	private int codigo;
	private String sigla;
	private HashSet<Alimentador> alimentadores;
	
	public Subestacao() {
		alimentadores = new HashSet<>();
	}
	
	public int getCodigo() {
		return codigo;
	}
	
	public void setCodigo(int codigo) {
		this.codigo = codigo;
	}
	
	public HashSet<Alimentador> getAlimentadores() {
		return alimentadores;
	}
	
	public void setAlimentadores(HashSet<Alimentador> alimentadores) {
		this.alimentadores = alimentadores;
	}

	public String getSigla() {
		return sigla;
	}

	public void setSigla(String sigla) {
		this.sigla = sigla;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + codigo;
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
		Subestacao other = (Subestacao) obj;
		if (codigo != other.codigo)
			return false;
		return true;
	}
	
}
