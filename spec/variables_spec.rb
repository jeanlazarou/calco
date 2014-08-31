require 'calco'

RSpec.configure do |c|
  c.alias_example_to :we_can
end

describe "With variables we can" do

  we_can "assign a value to a variable name" do

    doc = spreadsheet do

      definitions do
        set name: 'John'
      end
      
    end

    expect(doc.current[:name]).to eq('John')

  end

  we_can "change the value of a variable" do

    doc = spreadsheet do

      definitions do
        set name: 'John'
      end
      
    end

    sheet = doc.current

    sheet[:name] = 99

    expect(sheet[:name]).to eq(99)

  end
  
end
