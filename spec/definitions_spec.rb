require 'calco'

RSpec.configure do |c|
  c.alias_example_to :check
end

describe "definitions" do

  check "definitions setup" do
  
    doc = spreadsheet do
      
      definitions do
      
        set a: 8
        set b: 'hello'
        
        function f1: a
        function f2: b
      
      end
      
      sheet "A" do
        column value_of(:a)
        column value_of(:b)
        column :f1
        column :f2
      end
      
    end
  
  end

  check "definitions are visible to all sheets" do
  
    doc = spreadsheet do
      
      definitions do
      
        set a: 8
        
        function f1: a
      
      end
      
      sheet "A" do
        column value_of(:a)
        column :f1
      end
      
      sheet "B" do
        column value_of(:a)
        column :f1
      end
      
      sheet "C" do
        column value_of(:a)
        column :f1
      end
      
    end
  
  end

end
