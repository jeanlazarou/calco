require 'calco'

describe "Sheet selection" do

  it "a default sheet exists" do

    doc = spreadsheet do
    
    end

    expect(doc.current).to be_a(Calco::Sheet)

  end

  it "last sheet is the current one" do

    doc = spreadsheet do
    
      sheet "a" do
      end
      
      sheet "b" do
      end
      
    end

    expect(doc.current).to be(doc.sheet["b"])

  end

  it "selecting a sheet makes it current" do

    doc = spreadsheet do
    
      sheet "a" do
      end
      
      sheet "b" do
      end
      
    end

    doc.sheet("a")
    
    expect(doc.current).to eq(doc.sheet["a"])

  end

end
