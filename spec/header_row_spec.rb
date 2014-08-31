require 'calco'

describe 'Header row' do

  it "row[0] is an empty header/title row by default" do

    doc = spreadsheet do

      definitions do

        set price: 14.4

      end
      
      sheet do

        column value_of(:price)

      end

    end

    row = doc.row(0)

    expect(row).to be_empty

  end

  it "can be set with titles to columns" do

    doc = spreadsheet do

      definitions do

        set price: 14.4
        set product: 'USB 8Gb'
        set tax: 13
        set quantity: 30

        function total: price * quantity * (tax / 100 + 1)

      end
      
      sheet do

        column value_of(:product), title: 'Product'
        column value_of(:price), title: 'Price'
        column value_of(:quantity), title: 'Qty'
        column value_of(:tax), title: 'Tax rate'
        column skip
        column :total, title: 'Total'

      end

    end

    row = doc.row(0)

    expect(row).to match_array(['Product', 'Price', 'Qty', 'Tax rate', nil, 'Total'])

  end
  
end
