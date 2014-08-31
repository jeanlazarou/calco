require 'date'
require 'time'
require 'stringio'

require 'calco/engines/csv_engine'

module Calco

  describe CSVEngine do

    it "writes header when writing row 0" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('A').current

        sheet.write_row 0

      end

      expect(@result).to eq('c1,,c3,c4,c5')

    end

    it "writes empty values, numbers and functions" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('B').current

        sheet.write_row 1

      end

      expect(@result).to eq('"",=2*C1,11')

    end

    it "writes functions for times" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('time').current

        sheet.write_row 1

      end

      expect(@result).to eq('"=TIMEVALUE(""13:45:00"")"')

    end

    it "writes functions for dates" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('date').current

        sheet.write_row 1

      end

      expect(@result).to eq('"=DATEVALUE(""2013-07-27"")"')

    end

    it "writes functions for dates and style" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('date-style').current

        sheet.write_row 1

      end

      expect(@result).to eq('"=DATEVALUE(""2013-07-27"")+STYLE(IF(CURRENT()=TODAY();""red"";""default""))"')

    end

    it "writes functions for times and style" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('time-style').current

        sheet.write_row 1

      end

      expect(@result).to eq('"=TIMEVALUE(""13:45:00"")+STYLE(IF(CURRENT()=NOW();""red"";""default""))"')

    end

    it "writes functions for money" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('dollar').current

        sheet.write_row 1

      end

      expect(@result).to eq('=DOLLAR(11)')

    end

    it "writes money even for formulas" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('dollar-formulas').current

        sheet.write_row 1

      end

      expect(@result.split(',')).to eq(['=DOLLAR(11)', '=DOLLAR(2*A1)'])

    end

    it "surrounds formulas and style with money" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('dollar-conditional').current

        sheet.write_row 1

      end

      expect(@result).to eq('=DOLLAR(11),"=DOLLAR(2*A1+STYLE(IF(CURRENT()>2;""big"";""default"")))"')

    end

    it "writes % values" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('%').current

        sheet.write_row 1

      end

      expect(@result).to eq('=(0.3*100)%')

    end

    it "writes % functions" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('%-formulas').current

        sheet.write_row 1

      end

      expect(@result).to eq('11,=(2*A1)%')

    end

    it "writes sum aggregation" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('sum').current

        sheet.write_row 1
        sheet.write_row 2
        sheet.write_row 3

        sheet.replace_content :values, :total

        sheet.write_row 4

      end

      expect(@result.split).to match_array(['=DOLLAR(11)'] * 3 + ['=DOLLAR(SUM(A1:A3))'])

    end

    it "writes empty rows" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('B').current

        sheet.write_row 1
        sheet.empty_row
        sheet.write_row 3

      end

      expect(@result).to eq('"",=2*C1,11' "\n\n" '"",=2*C3,11')

    end

    it "supports constants in column layouts" do

      create_spreadsheet_and_save do |spreadsheet|

        sheet = spreadsheet.sheet('constants').current

        sheet.write_row 1

      end

      expect(@result).to eq('11,x 2,=2*A1')

    end

    before do
      @engine = Calco::CSVEngine.new
    end

    def create_spreadsheet_and_save &block

      @doc = spreadsheet(@engine) do

        definitions do

          set a: ''
          set b: 11
          set c: Time.parse('13:45')
          set d: Date.parse('2013-07-27')
          set e: 0.3

          function double: 2 * b

          function total: sum(b[(1..-1).as_grouping])

        end

        sheet "A" do

          column value_of(:a), :title => "c1"

          column :double

          column value_of(:b), :title => "c3"
          column value_of(:c), :title => "c4"
          column value_of(:d), :title => "c5"

        end

        sheet "B" do

          column value_of(:a)

          column :double

          column value_of(:b)

        end

        sheet "time" do
          column value_of(:c)
        end

        sheet "date" do
          column value_of(:d)
        end

        sheet "time-style" do
          column value_of(:c), style: _if(current == now(), 'red', 'default')
        end

        sheet "date-style" do
          column value_of(:d), style: _if(current == today(), 'red', 'default')
        end

        sheet "dollar" do
          column value_of(:b), :type => '$'
        end

        sheet "dollar-formulas" do
          column value_of(:b), :type => '$'
          column :double, :type => '$'
        end

        sheet "dollar-conditional" do
          column value_of(:b), :type => '$'
          column :double, :type => '$', style: _if(current > 2, 'big', 'default')
        end

        sheet "%" do
          column value_of(:e), :type => '%'
        end

        sheet "%-formulas" do
          column value_of(:b)
          column :double, :type => '%'
        end
        
        sheet "sum" do
          column value_of(:b), :type => '$', :id => :values
        end

        sheet "constants" do

          column value_of(:b)

          column "x 2"
          
          column :double

        end

      end

      buffer = StringIO.new

      @doc.save buffer, &block

      @result = buffer.string.strip

    end

  end

end
