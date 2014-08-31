require 'calco'
require 'calco/engines/csv_engine'

RSpec.configure do |c|
  c.alias_example_to :the
end

module Calco

  describe Spreadsheet do

    the 'Spreadsheet#row method returns value of current Sheet#row' do

      doc = create_sheet

      doc.sheet("sheet 1")

      expect(doc.row_to_hash(7)).to eq({
        'A' => '25',
        'B' => "2013-09-21",
        'C' => 'A7*7'
      })

    end

    the 'Spreadsheet#engine= method changes the engine' do

      doc = create_sheet

      doc.sheet("sheet 2")

      expect(doc.row_to_hash(7)).to eq({
        'A' => "2013-09-21",
        'B' => '2013-09-21; apply_style(IF(self=TODAY();"red";"default"))'
      })

      doc.engine = Calco::CSVEngine.new

      expect(doc.row_to_hash(7)).to eq({
        'A' => '=DATEVALUE("2013-09-21")',
        'B' => '=DATEVALUE("2013-09-21")+STYLE(IF(CURRENT()=TODAY();"red";"default"))'
      })

    end

    def create_sheet

      spreadsheet do

        definitions do

          set value: 25
          set some_date: "2013-09-21"

          function times_7: value * 7

        end

        sheet "sheet 1" do

          column value_of(:value)
          column value_of(:some_date)
          column :times_7

        end

        sheet "sheet 2" do

          column value_of(:some_date)
          column value_of(:some_date), style: _if(current == today(), 'red', 'default')

        end

      end

    end

  end

end
