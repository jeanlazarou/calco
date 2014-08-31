require 'calco'

#
# Next example shows different uses with formulas (functions) and styles.
#
# It 'implements' a calculator that evaluates which of the train or the car
# is better, depending on several parameters.
#
# It prints out the definition.
# It then prints out a CSV format after changing the engine.
#

require 'calco/engines/csv_engine'

doc = spreadsheet do

  definitions do

    set distance: 20 # km
    set fuel_price: 1.4 # €
    set consumption: 6 # l/100km
    set trip_duration: 50 # min

    set train_price: 8 # €
    set train_duration: 65 # min

    function trip_consumption: consumption * (distance / 100)

    function trip_price: consumption * fuel_price * (distance / 100)
    function benefit: (train_price - trip_price) / train_price

    function best_choice: _if(trip_duration < train_duration, 'car', 'train')

  end

  sheet do

    column value_of(:distance), :title => 'km'
    column value_of(:trip_duration), :title => 'min'
    column skip

    column value_of(:train_price), type: '$EUR', :title => 'train'
    column value_of(:train_duration), :title => 'min'
    column skip

    column :trip_consumption, :title => 'L'
    column :trip_price, style: _if(current > train_price, 'Red', 'default'), type: '$EUR', :title => 'price'
    column :benefit, :type => '%', :title => 'benefit'
    column skip

    column :best_choice, :title => 'choice'
    column skip

    column value_of(:fuel_price), type: '$EUR', :title => 'fuel price (1L)'
    column value_of(:consumption), :title => 'L/100km'

  end

end

puts "Layout definition:"

doc.current[:distance] = 40
cells = doc.row(3)

cells.each_index do |i|

  print "  #{Calco::COLUMNS[i]}: "

  cell = cells[i]

  puts cell

end

puts "\nCSV output:"

data = [
  {:distance => 40, :trip_duration => 30, :train_price => 10, :train_duration => 30, :fuel_price => 1.8, :consumption => 7},
  {:distance => 25, :trip_duration => 30, :train_price => 19, :train_duration => 70, :fuel_price => 1.3, :consumption => 5.4},
]

doc.engine = Calco::CSVEngine.new

doc.save($stdout) do |spreadsheet|

  sheet = doc.current

  sheet.write_row 0

  data.each_with_index do |entry, i|

    sheet.write_row i + 1, entry

  end

end
