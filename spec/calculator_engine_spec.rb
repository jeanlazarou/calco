require 'date'
require 'stringio'

require 'calco/engines/simple_calculator_engine'

module CalculatorSpecHelpers

  def compute n = nil
    
    buffer = StringIO.new
    
    @doc.save(buffer) do |spreadsheet|
      
      sheet = spreadsheet.current
      
      if n == 0
        
        sheet.write_row 0
        
      elsif n
      
        n.each {|i| sheet.write_row i}
      
      else
        
        sheet.write_row 1
        
      end
        
    end

    @result = buffer.string.split("\n").map {|s| s.strip}
    
  end
  
end

module Calco

  describe SimpleCalculatorEngine do
    
    include CalculatorSpecHelpers
  
    it "does not return header" do

      create_spreadsheet :default

      compute 0

      @result.should be_empty
      
    end
    
    it "computes simple function" do

      create_spreadsheet :default

      compute

      @result.last.should == "C1 = 20"
      
    end

    it "uses variable names as names" do

      create_spreadsheet :default

      compute
      
      @result == ["price = 10", "quantity = 2", "C0 = 20"]
      
    end

    it "uses titles as names" do

      create_spreadsheet :with_title

      compute

      @result.should == ["price = 10", "quantity = 2", "P*Q = 20"]
      
    end

    it "computes sheets with absolute cell references" do

      create_spreadsheet :with_title, :add_absolute

      compute 1..3

      @result[3].should == "tax = 100"
      
      @result[11].should == "price = 10"
      @result[12].should == "quantity = 2"
      @result[13].should == "P*Q = 20"
      @result[14].should == "F3 = 120"
      
    end

    it "uses string values" do

      @doc = spreadsheet(@engine) do

        definitions do

          set first_name: "John"
          set last_name: "Smith"
          
          function name: first_name + " " + last_name
          
        end
        
        sheet do

          column value_of(:first_name)
          column value_of(:last_name)
          
          column :name, title: 'name'

        end

      end

      compute
      
      @result.last.should == 'name = "John Smith"'
      
    end
    
    before do
      @engine = Calco::SimpleCalculatorEngine.new
    end

    def create_spreadsheet which, add_absolute_row = false

      @doc = spreadsheet(@engine) do

        definitions do

          set price: 10
          set quantity: 2
          
          function total: price * quantity
          
          if add_absolute_row == :add_absolute
          
            set tax: 100
            
            function more: total + tax
            
          end
          
        end
        
        sheet do

          column value_of(:price)
          column value_of(:quantity)
          
          column :total                if which == :default
          column :total, title: 'P*Q'  if which == :with_title

          if add_absolute_row == :add_absolute
            
            column skip
            
            column value_of(:tax, :absolute)
            column :more
            
            tax.absolute_row = 1
            
          end
          
        end

      end

    end
    
  end

  describe SimpleCalculatorEngine do
    
    include CalculatorSpecHelpers
    
    it "computes functions with today and year" do

      @doc = spreadsheet(@engine) do

        definitions do

          set some_date: "2013/09/06"

          function some_year: year(some_date)
          function today: year(today)

        end
        
        sheet do
        
          column value_of(:some_date)
          
          column :some_year
          column :today
          
        end

      end

      compute

      year = Date.today.year
      
      @result.should == ['some_date = "2013-09-06"', "B1 = 2013", "C1 = #{year}"]
      
    end
    
    it "computes functions with left" do

      @doc = spreadsheet(@engine) do

        definitions do

          set name: "James"

          function initials: left(name, 1)

        end
        
        sheet do
        
          column value_of(:name)
          
          column :initials
          
        end

      end

      compute

      @result.should == ['name = "James"', 'B1 = "J"']
      
    end
    
    before do
      @engine = Calco::SimpleCalculatorEngine.new
    end

  end
  
end
