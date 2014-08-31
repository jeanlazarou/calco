require 'calco'

describe "Spreadsheet with conditions" do

  it "supports conditional functions" do

    doc = spreadsheet do

      definitions do

        set score: 23
        function opinion: _if(score < 20, 'Oops...', 'Great!')
        
      end
      
      sheet do

        column value_of(:score)
        column :opinion

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('23')
    expect(row[1]).to eq('IF(A1<20;"Oops...";"Great!")')

  end

  it "supports conditions with strings" do

    doc = spreadsheet do

      definitions do

        set who: "John"
        function bonus: _if(who == 'John', 200, 0)
        
      end
      
      sheet do

        column value_of(:who)
        column :bonus

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('IF(A1="John";200;0)')

  end

  it "accepts conditions within conditions" do

    doc = spreadsheet do

      definitions do

        set age: 34
        set who: "John"

        function bonus: _if(who == 'John', _if(age > 25, 200, 0), 10)

      end
      
      sheet do
      
        column value_of(:who)
        column value_of(:age)

        column :bonus

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('34')
    expect(row[2]).to eq('IF(A1="John";IF(B1>25;200;0);10)')

  end

  it "supports conditions with or" do

    doc = spreadsheet do

      definitions do

        set who: "John"
        function bonus: _if(_or(who == 'John', who == 'Mary'), 200, 0)

      end
      
      sheet do

        column value_of(:who)
        column :bonus

      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('"John"')
    expect(row[1]).to eq('IF(OR((A1="John");(A1="Mary"));200;0)')

  end

end
