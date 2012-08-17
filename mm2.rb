require 'watir'
require 'nokogiri'
require 'win32ole'

#loop through the number of rows in my excel sheet

#open a browser and go to an address
#make title optional; specify which sheet you want to go to
def browser
	# could make it #=> def ie_address(address, title)
	@browser = Watir::Browser.new
	# @browser = Watir::IE.attach(:title, title)
end

#open an excel file, select a sheet number; could make the sheet number optional, need a condition in the code
def xl_open(my_dir, sheet_num)
	excel = WIN32OLE.new('Excel.Application')
	excel.visible = true
	@workbook = excel.Workbooks.Open(my_dir);
	@worksheet = @workbook.Worksheets(sheet_num)
end

# put down different headers for each sheet number
def header(row, column, title)
	@worksheet.setproperty('Cells',row,column, title)
end

#parse the page using nokogiri, loop through the selectors information and put it into an excel column
# double check if you want to loop here, it's possible you don't want to.. 
def scrape2xl(selector, address, start_row, start_column, direction="horizontal")
	@browser.goto address
	@page_html = Nokogiri::HTML.parse(@browser.html)
	items = @page_html.search(selector).map(&:text).map(&:strip)
	items.each do |x|
		@worksheet.setproperty('Cells', start_row, start_column, x)
		if direction == "column" 	
			start_column += 1
		elsif direction == "horizontal"
			start_row += 1
		end
	end
end

# |   a;1    |   b;2   |  c;3 |  d;4   |  e;5 |    
# | Selector | Website | Name |	Source | Pull |

xl_open("#{Dir.pwd}/mm2.xls", 1)
browser

n = 2
loop until n > @worksheet.Cells(1).value
scrape2xl(@worksheet.Cells(n,1).value, @worksheet.Cells(n,2).value, n, 5)
	n += 1
	break if n > @worksheet.Cells(1,1).value 
end

@browser.close




