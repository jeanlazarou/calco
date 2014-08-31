require 'calco'

describe 'Spreadsheet built-in functions' do

  it "supports 'LEFT'" do

    doc = spreadsheet do

      definitions do

        set name: "Miles"

        function initial: left(name, 1)

      end

      sheet do

        column value_of(:name), title: 'Name'
        column :initial, title: 'Initial'

      end

    end

    row = doc.row(1)

    expect(row[1]).to eq('LEFT(A1, 1)')

  end

  it "supports 'TODAY' and 'YEAR'" do

    doc = spreadsheet do

      definitions do

        set birth_date: "2004-09-18"

        function age: year(today()) - year(birth_date)

      end

      sheet do

        column value_of(:birth_date), title: 'Birth date'
        column :age, title: 'Age'

      end

    end

    row = doc.row(1)

    expect(row[1]).to eq('YEAR(TODAY())-YEAR(A1)')

  end

  it "registers new functions" do

    Calco::BuiltinFunction.declare :my_func, 0, Integer

    doc = spreadsheet do

      definitions do
        function x: my_func
      end

      sheet do
        column :x
      end

    end

    row = doc.row(1)

    expect(row[0]).to eq('MY_FUNC()')

  end

  it "complains if function does not exist" do

    expect {

      spreadsheet do

        definitions do
          function x: my_func(13)
        end

      end

    }.to raise_error("Unknown function or variable 'my_func'")


  end

  it "complains if wrong number of arguments" do

    Calco::BuiltinFunction.declare :my_func, 0, Integer

    expect {

      spreadsheet do

        definitions do

          function x: my_func(13)

        end

      end

    }.to raise_error(ArgumentError, "Function MY_FUNC requires 0, was 1 ([13])")


  end

  it "complains if missing arguments" do

    Calco::BuiltinFunction.declare :my_func, 2, Integer

    expect {

      spreadsheet do

        definitions do

          function x: my_func(13)

        end

      end

    }.to raise_error(ArgumentError, "Function MY_FUNC requires 2, was 1 ([13])")


  end

  it "accepts variable arity for functions" do

    Calco::BuiltinFunction.declare :my_func, :n, Integer

    spreadsheet do

      definitions do

        function x: my_func(13, 12, 3)
        function y: my_func

      end

    end

  end

  after do
    Calco::BuiltinFunction.undeclare :my_func
  end

end
