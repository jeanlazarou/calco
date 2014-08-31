require 'calco'

describe "Spreadsheet's sheet" do

  it "supports dynamic styles" do

    doc = spreadsheet do

      definitions do

        set price: 14.4

      end
      
      sheet do

        column value_of(:price), style: _if(current > 120, 'Alert', 'default')

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('14.4; apply_style(IF(self>120;"Alert";"default"))')

  end

end
