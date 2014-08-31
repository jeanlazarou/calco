require 'calco'

describe "Implicit type conversion" do

  it "detects time values" do
  
    doc = spreadsheet do
      
      definitions do
      
        set time: '8:11'
        set not_time: '13,01'
      
      end
      
    end
  
    expect(doc.current[:time]).to be_a(Time)
    expect(doc.current[:not_time]).to be_a(String)
    
  end

  it "detects date values" do
  
    doc = spreadsheet do
      
      definitions do
      
        set date1: '2013-06-17'
        set date2: '2013/06/17'
        set not_date: '2013 06 17'
      
      end
      
    end
  
    expect(doc.current[:date1]).to be_a(Date)
    expect(doc.current[:date2]).to be_a(Date)
    expect(doc.current[:not_date]).to be_a(String)
    
  end

end
