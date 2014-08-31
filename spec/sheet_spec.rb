require 'stringio'

require 'calco'

describe "Sheet" do

  it "produces one cell with some assigned literal" do

    doc = spreadsheet do
      sheet do
        column 77
      end
    end

    expect(doc.row(1)).to include('77')

  end

  it "produces one cell with the value of a variable" do

    doc = spreadsheet do

      definitions do

        set name: 'John'
        
      end
      
      sheet do

        column value_of(:name)

      end

    end

    expect(doc.row(1)).to include('"John"')

  end

  it "produces a reference to a variable" do

    doc = spreadsheet do

      definitions do

        set name: 'John'
        function john: name
        
      end
      
      sheet do

        column value_of(:name)
        column :john

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('A1')

  end

  it "produces references with row index" do

    doc = spreadsheet do

      definitions do

        set name: 'John'
        function john: name
        
      end
      
      sheet do

        column value_of(:name)
        column :john

      end

    end

    sheet = doc.current

    sheet[:name] = 'James'

    row = sheet.row(4)

    expect(row[0]).to eq('"James"')
    expect(row[1]).to eq('A4')

  end

  it "accepts references to higher cells" do

    doc = spreadsheet do

      definitions do

        set name: 'John'
        function john: name
        
      end
      
      sheet do

        column :john
        column value_of(:name)

      end

    end

    sheet = doc.current

    sheet[:name] = 'James'

    row = sheet.row(2)

    expect(row[0]).to eq('B2')
    expect(row[1]).to eq('"James"')

  end

  it "supports blank cells" do

    doc = spreadsheet do

      definitions do

        set name: 'John'
        function john: name
        
      end
      
      sheet do
      
        column value_of(:name)
        column skip
        column :john

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('')
    expect(row[2]).to eq('A1')

  end

  it "implements #record_assign" do

    doc = spreadsheet do

      definitions do

        set name: ''
        set last_name: ''
        
      end
      
      sheet do
      
        column value_of(:name)
        column value_of(:last_name)

      end

    end

    sheet = doc.current

    sheet.record_assign :name => "John", :last_name => "Smith"
    
    row = doc.row(1)
    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('"Smith"')

    sheet.record_assign :name => "James", :last_name => "Black"
    
    row = doc.row(2)
    expect(row[0]).to eq('"James"')
    expect(row[1]).to eq('"Black"')

  end

  it "implements #write_row with data" do

    doc = spreadsheet do

      definitions do

        set name: ''
        set last_name: ''
        
      end
      
      sheet do
      
        column value_of(:name)
        column value_of(:last_name)

      end

    end

    out = StringIO.new
    
    doc.save out do |spreadsheet|
    
      sheet = spreadsheet.current

      sheet.write_row 1, :name => "John", :last_name => "Smith"
      
    end

    expect(out.string).to eq(%|A1: "John"\nB1: "Smith"\n\n|)
    
  end

end
