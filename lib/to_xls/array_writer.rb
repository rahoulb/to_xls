module ToXls

  class ArrayWriter
    def initialize(array, options = {})
      @array = array
      @options = options
    end

    def write_string(string = '')
      io = StringIO.new(string)
      write_io(io)
      io.string
    end
    
    def write_sheet(sheet)
      DataWriter.new(@array, @options).write_to(sheet)
    end

    def write_book(book)
      sheet = DataSheet.new(book, @options[:name] || 'Sheet 1', @options)
      sheet.write(@array)
    end

    def write_io(io)
      book = Spreadsheet::Workbook.new
      write_book(book)
      book.write(io)
    end
  end

end
