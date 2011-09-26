module ToXls

  class DataSheet

    def initialize(book, title, options = {})
      @book = book
      @title = title
      @options = options
    end
    
    def write(array)
      binder = DataWriter.new(array, @options)
      binder.write_to(sheet)
    end

    def sheet
      @sheet ||= begin
        sheet = nil
        @book.worksheets.each do |worksheet|
          if worksheet.name == @title
            sheet = worksheet
            break
          end
        end
        sheet ||= begin
          sheet = @book.create_worksheet
          sheet.name = @title
          sheet
        end
        sheet
      end
    end
    
  end

end

