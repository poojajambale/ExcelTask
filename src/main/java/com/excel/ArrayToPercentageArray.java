package com.excel;

public class ArrayToPercentageArray {

	public static void main(String[] args) {
		
//		int[] arr = new int[set.size()];
//		int[] arr = new int[10];
		int[] arr = {1};
		double total = 0;
		double percentage;
		
		for (int count1 : arr) {
//			System.out.println(count1);
			total += count1;
		}
		
		System.out.println("total:"+total);
		
		for (int count2 : arr) {
			System.out.println("count2:"+count2);
			percentage = (count2/total)*100;
			System.out.println("percentage:"+String.format("%.2f", percentage));
			
		}
	} 
}
