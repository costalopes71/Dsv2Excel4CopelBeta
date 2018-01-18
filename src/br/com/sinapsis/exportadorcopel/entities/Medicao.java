package br.com.sinapsis.exportadorcopel.entities;

import java.util.Calendar;

public class Medicao {

	private Calendar data;
	private int corrFaseA;
	private int corrFaseB;
	private int corrFaseC;
	private int potAtiva;
	private int potReat;
	private double fatorPot;
	private double tensaoA;
	private double tensaoB;
	private double tensaoC;
	
	public Calendar getData() {
		return data;
	}
	
	public void setData(Calendar data) {
		this.data = data;
	}
	
	public int getCorrFaseA() {
		return corrFaseA;
	}
	
	public void setCorrFaseA(int corrFaseA) {
		this.corrFaseA = corrFaseA;
	}
	
	public int getCorrFaseB() {
		return corrFaseB;
	}
	
	public void setCorrFaseB(int corrFaseB) {
		this.corrFaseB = corrFaseB;
	}
	
	public int getCorrFaseC() {
		return corrFaseC;
	}
	
	public void setCorrFaseC(int corrFaseC) {
		this.corrFaseC = corrFaseC;
	}
	
	public int getPotAtiva() {
		return potAtiva;
	}
	
	public void setPotAtiva(int potAtiva) {
		this.potAtiva = potAtiva;
	}
	
	public int getPotReat() {
		return potReat;
	}
	
	public void setPotReat(int potReat) {
		this.potReat = potReat;
	}
	
	public double getFatorPot() {
		return fatorPot;
	}
	
	public void setFatorPot(double fatorPot) {
		this.fatorPot = fatorPot;
	}
	
	public double getTensaoA() {
		return tensaoA;
	}

	public void setTensaoA(double tesaoA) {
		this.tensaoA = tesaoA;
	}
	
	public double getTensaoB() {
		return tensaoB;
	}
	
	public void setTensaoB(double tesaoB) {
		this.tensaoB = tesaoB;
	}
	
	public double getTensaoC() {
		return tensaoC;
	}

	public void setTensaoC(double tesaoC) {
		this.tensaoC = tesaoC;
	}
	
}
