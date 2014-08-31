
require 'calco'

describe 'Absolute references' do

  it "supports absolute references" do

    doc = spreadsheet do

      definitions do

        set price: 14.4
        set tax: 13

        function total: price * (tax / 100 + 1)

      end
      
      sheet do

        column value_of(:price)
        column skip
        column :total

        column skip
        column skip

        column value_of(:tax, :absolute)

        tax.absolute_row = 1

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('14.4')
    
    expect(row[2]).to eq('A1*(($F$1/100)+1)')
    
    expect(row[5]).to eq('13')

  end

  it "only sets cell content for the absolute row" do

    doc = spreadsheet do

      definitions do

        set tax: 13
        set some_val: 77

      end
      
      sheet do

        column value_of(:tax, :absolute)
        column value_of(:some_val)
        column :some_val

        tax.absolute_row = 7

      end

    end

    row = doc.row(1)
    expect(row[0]).to be_nil
    expect(row[1]).to eq('77')
    expect(row[2]).to eq("C1")
    
    row = doc.row(7)
    expect(row[0]).to eq('13')
    expect(row[1]).to eq('77')
    expect(row[2]).to eq("C7")
    
    row = doc.row(8)
    expect(row[0]).to be_nil
    expect(row[1]).to eq('77')
    expect(row[2]).to eq("C8")

  end

end
