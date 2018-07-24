package com.excel1;

public class Book {

	private String title;
	private String author;
	private float price;
	public Book() {
}
	
	public String toString() 
	{	
	return String.format("%s - %s - %f",title,author,price);
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public float getPrice() {
		return price;
	}

	public void setPrice(double d) {
		this.price = (float) d;
	}
	
	
}
