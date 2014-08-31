require 'calco'

RSpec.configure do |c|
  c.alias_example_to :support
end

describe "Functions support" do

  support "operators" do

    doc = spreadsheet do

      definitions do

        set price: 14.4
        set product: 'USB 8Gb'
        set tax: 13
        set quantity: 30

        function total: price * quantity * (tax / 100 + 1)

      end
      
      sheet do

        column value_of(:product)
        column value_of(:price)
        column value_of(:quantity)
        column value_of(:tax)
        column skip
        column :total

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"USB 8Gb"')
    expect(row[1]).to eq('14.4')
    expect(row[2]).to eq('30')
    expect(row[3]).to eq('13')
    
    expect(row[5]).to eq('(B1*C1)*((D1/100)+1)')

  end

  support "functions using functions" do

    doc = spreadsheet do

      definitions do

        set price: 14.4
        set product: 'USB 8Gb'
        set tax: 13
        set quantity: 30

        function total: price * quantity
        function total_with_taxes: total * (tax / 100 + 1)

      end
      
      sheet do

        column value_of(:product)
        column value_of(:price)
        column value_of(:quantity)
        column value_of(:tax)
        column skip
        column :total
        column :total_with_taxes

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"USB 8Gb"')
    expect(row[1]).to eq('14.4')
    expect(row[2]).to eq('30')
    expect(row[3]).to eq('13')
    
    expect(row[5]).to eq('B1*C1')
    expect(row[6]).to eq('F1*((D1/100)+1)')

  end

  support "conditions with functions" do

    doc = spreadsheet do

      definitions do

        set x: 9

        function f0: x + 3
        function f1: f0 * f0
        function f2: f1 + f0
        function f3: f2 + 7

        function cond: _if(f1 > 2, f2, f3)

      end
      
      sheet do
      
        column value_of(:x)
        column :f0
        column :f1
        column :f2
        column :f3

        column :cond

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('9')
    expect(row[1]).to eq('A1+3')
    expect(row[2]).to eq('B1*B1')
    expect(row[3]).to eq('C1+B1')
    expect(row[4]).to eq('D1+7')
    expect(row[5]).to eq('IF(C1>2;D1;E1)')

  end

  support "containing literals" do

    some_day = Date.new(1998, 7, 12)
    
    doc = spreadsheet do

      definitions do
        
        function a: 'string'
        function b: 12
        function c: 34.9
        function d: some_day
        
      end
      
      sheet do
        
        column :a
        column :b
        column :c
        column :d
        
      end

    end

    row = doc.current.row(1)

    expect(row[0]).to eq('"string"')
    expect(row[1]).to eq("12")
    expect(row[2]).to eq("34.9")
    expect(row[3]).to eq(some_day.to_s)

  end

  support "starting with an integer" do

    doc = spreadsheet do

      definitions do

        set x: 12
        
        function add: 66 + x

      end
      
      sheet do

        column value_of(:x)
        column :add

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('12')
    expect(row[1]).to eq('66+A1')

  end

  support "starting with a float" do

    doc = spreadsheet do

      definitions do

        set x: 12.9
        
        function add: 34.7 + x

      end
      
      sheet do

        column value_of(:x)
        column :add

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('12.9')
    expect(row[1]).to eq('34.7+A1')

  end

  support "starting with a string" do

    doc = spreadsheet do

      definitions do

        set name: "Joe"
        
        function greeting: "Hello " + name

      end
      
      sheet do

        column value_of(:name)
        column :greeting

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"Joe"')
    expect(row[1]).to eq('"Hello "+A1')

  end
  
end
