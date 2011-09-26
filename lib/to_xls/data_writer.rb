module ToXls

  class DataWriter

    def initialize(array, options = {})
      @array = array
      @options = options
    end

    def columns
      return  @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError.new(":columns (#{columns}) must be an array or nil") unless (@columns.nil? || @columns.is_a?(Array))
      @columns ||= can_get_columns_from_first_element? ? get_columns_from_first_element : []
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers] || columns
      raise ArgumentError, ":headers (#{@headers.inspect}) must be an array" unless @headers.is_a? Array
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end
    
    def write_to(sheet)
      if columns.any?
        if headers_should_be_included? and sheet.row_count == 0
          fill_row(sheet, sheet.row(0), headers)
        end

        @array.each do |model|
          fill_row(sheet, sheet.row(sheet.row_count), columns, model.respond_to?(:as_xls) ? model.as_xls(@options): model)
        end
      end
    end
    
    private

      def can_get_columns_from_first_element?
        @array.first &&
        @array.first.respond_to?(:as_xls) || (
          @array.first.respond_to?(:attributes) &&
          @array.first.attributes.respond_to?(:keys) &&
          @array.first.attributes.keys.is_a?(Array)
        )
      end

      def get_columns_from_first_element
        @array.first.respond_to?(:as_xls) ? 
          @array.first.as_xls(@options).keys : 
          @array.first.attributes.keys.sort_by {|sym| sym.to_s}.collect.to_a
      end

      def fill_row(sheet, row, column, model=nil)
        case column
        when String, Symbol
          data = model ? (model.is_a?(Hash) ? model[column] : model.send(column)) : column
          if data.is_a?(Enumerable)
            nested_options = @options.dup
            nested_options.delete(:headers) if @options[:headers].is_a?(Array)
            nested_options.delete(:columns) if @options[:columns].is_a?(Array)
            sheet = DataSheet.new(sheet.workbook, column, nested_options)
            sheet.write(data)
            data = data.count
          end
          row.push(data)
        when Hash
          column.each{|key, values| fill_row(sheet, row, values, model && (model.is_a?(Hash) ? model[key] : model.send(key)))}
        when Array
          column.each{|value| fill_row(sheet, row, value, model)}
        else
          raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
        end
      end

  end

end

