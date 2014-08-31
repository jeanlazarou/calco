require 'calco'

module Calco
  
  describe DefaultEngine do
  
    it 'generates value of var for "column value_of(:var)"' do

      doc = spreadsheet do

        definitions do

          set a: "12:58"
          
        end
        
        sheet do

          column value_of(:a)
          
        end
        
      end

      row = doc.current.row(1)

      expect(row[0]).to eq('12:58:00')
      
    end
    
    it 'generates (invalid) reference to same cell for "column :var"' do

      doc = spreadsheet do

        definitions do

          set a: "12:58"
          
        end
        
        sheet do

          column :a
          
        end
        
      end

      row = doc.current.row(1)

      expect(row[0]).to eq('A1')
      
    end
    
    it 'generates "" for "column skip"' do

      doc = spreadsheet do

        definitions do

          set a: "12:58"
          
        end
        
        sheet do

          column skip
          column value_of(:a)
          
        end
        
      end

      row = doc.current.row(1)

      expect(row[0]).to eq('')
      expect(row[1]).to eq('12:58:00')
      
    end
    
    it "generates all values as string literals" do

      doc = spreadsheet do

        definitions do

          set a: 'Hello'
          set b: 12
          
          function c: a + "Hi"
          function d: b + 123
          function e: 76.9
          function f: "Max"
        
          set g: "2013-09-29"
          set h: "12:58"
          
        end
        
        sheet do

          column value_of(:a)
          column value_of(:b)
          
          column :c
          column :d
          column :e
          column :f

          column value_of(:g)
          column value_of(:h)
          
          column :h
          
        end
        
      end

      row = doc.current.row(1)

      expect(row[0]).to eq('"Hello"')
      expect(row[1]).to eq('12')
      expect(row[2]).to eq('A1+"Hi"')
      expect(row[3]).to eq('B1+123')
      expect(row[4]).to eq('76.9')
      expect(row[5]).to eq('"Max"')
      expect(row[6]).to eq('2013-09-29')
      expect(row[7]).to eq('12:58:00')
      expect(row[8]).to eq('I1')
      
    end
  
  end
  
end
