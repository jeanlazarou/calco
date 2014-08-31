require 'calco/engines/default_engine'

describe 'Spreadsheet' do

  it "accepts column range" do

    doc = spreadsheet do

      definitions do

        set quantity: 17
        function total: sum(quantity[1..-1])

      end

      sheet do

        column value_of(:quantity)
        column :total

        total.absolute_row = 1

      end

    end

    sheet = doc.current

    row = sheet.row(1)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq("SUM(A1:A#{Calco::DefaultEngine.row_infinity})")

    row = sheet.row(10)
    expect(row[0]).to eq('17')
    expect(row[1]).to be_nil

  end

  it "accepts range linked to variable" do

    doc = spreadsheet do

      definitions do

        set quantity: 17
        function total: sum(quantity[10..21])

      end

      sheet do

        column value_of(:quantity)
        column :total

        total.absolute_row = 1

      end

    end

    sheet = doc.current

    row = sheet.row(1)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq("SUM(A10:A21)")

    row = sheet.row(10)
    expect(row[0]).to eq('17')
    expect(row[1]).to be_nil

  end

  it "accepts aggregate in same column with function" do

    doc = spreadsheet do

      definitions do

        set unit_price: 33
        set quantity: 17

        function price: unit_price * quantity
        function total: sum(price[(1..-1).as_grouping])

      end

      sheet do

        column value_of(:unit_price), :id => :up
        column value_of(:quantity), :id => :q
        column :price, :id => :p

      end

    end

    sheet = doc.current

    5.times {|i| sheet.row(i);}

    row = sheet.row(6)
    expect(row[0]).to eq('33')
    expect(row[1]).to eq('17')
    expect(row[2]).to eq('A6*B6')

    sheet.replace_content :up, nil
    sheet.replace_content :q, nil
    sheet.replace_content :p, :total

    row = sheet.row(7)
    expect(row[2]).to eq("SUM(C1:C6)")

    expect(row[0]).to eq('')
    expect(row[1]).to eq('')

  end

  it "accepts aggregate in same column" do

    doc = spreadsheet do

      definitions do

        set quantity: 17

        function sub_total: sum(quantity[(1..-1).as_grouping])

      end

      sheet do

        column value_of(:quantity), :id => :here

      end

    end

    sheet = doc.current

    row = sheet.row(10)
    expect(row[0]).to eq('17')

    sheet.replace_content :here, :sub_total

    row = sheet.row(11)
    expect(row[0]).to eq("SUM(A1:A10)")

    sheet.restore_content

    row = sheet.row(12)
    expect(row[0]).to eq('17')

  end

  it "accepts range reference to following column" do

    doc = spreadsheet do

      definitions do

        set quantity: 17
        function total: sum(quantity[1..-1])

      end

      sheet do

        column :total
        column value_of(:quantity)

        total.absolute_row = 1

      end

    end

    sheet = doc.current

    row = sheet.row(1)
    expect(row[0]).to eq("SUM(B1:B#{Calco::DefaultEngine.row_infinity})")
    expect(row[1]).to eq('17')

    row = sheet.row(10)
    expect(row[0]).to be_nil
    expect(row[1]).to eq('17')

  end

end
