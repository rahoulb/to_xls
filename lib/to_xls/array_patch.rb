class Array
  # Options for to_xls: columns, name, header, sheet
  def to_xls(options = {})
    ToXls::ArrayWriter.new(self, options).write_string
  end
end
