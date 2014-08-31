require 'calco'

RSpec.configure do |c|
  c.alias_example_to :change_by
end

describe 'Content change' do

  change_by "replacing content with a function" do

    doc = spreadsheet do

      definitions do
        
        set empty: ''
        set quantity: 17
        
        function double: quantity * 2

      end
      
      sheet do
        
        column value_of(:quantity)
        column value_of(:empty), :id => :here
        
      end

    end

    sheet = doc.current
    
    sheet.replace_content :here, :double
    
    row = sheet.row(2)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq("A2*2")
    
  end
  
  change_by "replacing content with a variable" do

    doc = spreadsheet do

      definitions do
        
        set yes: 'Yes'
        set empty: ''
        set quantity: 17
        
      end
      
      sheet do
        
        column value_of(:quantity)
        column value_of(:empty), :id => :here
        
      end

    end

    sheet = doc.current
    
    sheet.replace_content :here, :yes
    
    row = sheet.row(2)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq('"Yes"')
    
  end
  
  change_by "replacing content with a value" do

    doc = spreadsheet do

      definitions do
        
        set yes: 'Yes'
        set empty: ''
        set quantity: 17
        
      end
      
      sheet do
        
        column value_of(:quantity)
        column value_of(:empty), :id => :here
        
      end

    end

    sheet = doc.current
    
    sheet.replace_content :here, 32
    
    row = sheet.row(2)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq("32")
    
  end
  
  change_by "changing and resetting content" do

    doc = spreadsheet do

      definitions do
        
        set empty: ''
        set quantity: 17
        
        function double: quantity * 2
        function triple: quantity * 3

      end
      
      sheet do
        
        column value_of(:quantity)
        column value_of(:empty), :id => :here
        column value_of(:empty), :id => :here_also
        
      end

    end

    sheet = doc.current
    
    row = sheet.row(1)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq('""')
    expect(row[2]).to eq('""')
    
    sheet.replace_content :here, :double
    sheet.replace_content :here_also, :triple
    
    row = sheet.row(2)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq("A2*2")
    expect(row[2]).to eq("A2*3")

    sheet.restore_content
    
    sheet[:empty] = "Yes"
    
    row = sheet.row(3)
    expect(row[0]).to eq('17')
    expect(row[1]).to eq('"Yes"')
    expect(row[2]).to eq('"Yes"')
    
  end

  change_by "mass-replacing content with values and nil values" do

    doc = spreadsheet do

      definitions do
        
        set yes: 'Yes'
        set empty: ''
        set quantity: 17
        
      end
      
      sheet do
        
        column value_of(:quantity)
        column value_of(:empty), :id => :one
        column value_of(:yes)
        column value_of(:empty), :id => :two
        column value_of(:quantity), :id => :three
        
      end

    end

    sheet = doc.current
    
    sheet.replace_and_clear :one => 19, :three => 7
    
    row = sheet.row(2)
    expect(row[0]).to eq('')
    expect(row[1]).to eq("19")
    expect(row[2]).to eq('')
    expect(row[3]).to eq('')
    expect(row[4]).to eq('7')
    
  end
  
end
