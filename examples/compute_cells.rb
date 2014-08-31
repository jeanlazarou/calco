require 'calco/engines/simple_calculator_engine'

#
# Next example uses a SimpleCaclulatorEngine that does not produce a spreadsheet
# with formulas but computes the calculations instead.
#

engine = Calco::SimpleCalculatorEngine.new

doc = spreadsheet(engine) do

  definitions do

    set tax_rate: 0
    
    set basic_price: 0
    set quantity: 1
    
    function price: (basic_price * quantity) * (1 + (tax_rate / 100.0))
    
  end
  
  sheet do

    column value_of(:tax_rate, :absolute)

    column value_of(:basic_price)
    column value_of(:quantity)
    
    column skip
    
    column :price, title: 'price'

    tax_rate.absolute_row = 1
    
  end
  
end

data = [
  [12, 3],
  [10.3, 2],
  [34, 4]
]

doc.save($stdout) do |spreadsheet|
  
  sheet = spreadsheet.current
  
  sheet[:tax_rate] = 100
  
  data.each_with_index do |item, i|
    
    sheet[:basic_price] = item[0]
    sheet[:quantity] = item[1]
    
    sheet.write_row i + 1
    
  end
  
end
