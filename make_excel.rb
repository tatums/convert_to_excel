require 'csv'
require 'spreadsheet'

@file = File.open("CUST_PCS_INV.csv", "r")

Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet  

#FORMATING
bold = Spreadsheet::Format.new({:weight => :bold, :size => 10} )
sheet1.column(0).width = 15
sheet1.column(1).width = 15
sheet1.column(2).width = 15

sheet1.row(0).default_format = bold
sheet1[0,0] = 'CUSTOMER' 
sheet1[0,1] = 'PIECE COUNT' 
sheet1[0,2] = 'ORDER' 


i = 1

CSV.parse(@file, {:headers =>true, :col_sep => "\t" }) do |row|     
     
  sheet1.row(i).concat [ row[0],row[1], row[2]]
  i += 1
     
end
book.write "CUST_PCS_INV.xls"


