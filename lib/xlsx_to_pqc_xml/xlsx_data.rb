require 'rubyXL'
module XlsxToPqcXml
  class XlsxData

    attr_reader :xlsx_path

    # TODO: Validate required values present
    # TODO: Add value splitting

    ##
    # Create a new XlsxData for XLSX file `xlsx_path` with `config` hash.
    #
    # Config hash is similar to the following:
    #
    #     {:sheet_name=>"Structural",
    #      :sheet_position=>0,
    #      :heading_type=>"row",
    #      :attributes=>
    #        [
    #         { :attr=>"ark_id",
    #           :headings=>["ARK ID"],
    #           :requirement=>"required"},
    #         { :attr=>"page_sequence",
    #           :headings=>["PAGE SEQUENCE"],
    #           :requirement=>"required"},
    #         { :attr=>"filename",
    #           :headings=>["FILENAME"],
    #           :requirement=>:required},
    #         { :attr=>"visible_page",
    #           :headings=>["VISIBLE PAGE"],
    #           :requirement=>"required"},
    #         { :attr=>"toc_entry",
    #           :headings=>["TOC ENTRY"]},
    #         { :attr=>"ill_entry",
    #           :headings=>["ILL ENTRY"]},
    #         { :attr=>"notes",
    #           :headings=>["NOTES"]}
    #        ]
    #      }
    #
    # @param [String] xlsx_path path to the XLSX file
    # @param [Hash] config spreadsheet configuration
    def initialize xlsx_path:, config: {}
      @xlsx_path        = xlsx_path
      @sheet_config     = config
      @data             = nil
      @headers          = []
    end

    ##
    # @return [Array<Hash>] array of the spreadsheet data as hashes
    # @raise [StandardError] if there are duplicate headers or missing columns
    def data
      return @data unless @data.nil?

      @data = []

      xlsx      = RubyXL::Parser.parse xlsx_path
      worksheet = xlsx[0]

      validate_headers

      if @sheet_config.fetch(:heading_type, :row).to_sym == :column
        # headings are in the first column; for each header, work across the
        # row, collecting the value in each column.
        headers.each_with_index do |head, row_pos|
          next if head.nil?
          worksheet.sheet_data.rows[row_pos].cells.each_with_index do |cell,col_pos|
            # don't process the first column; it has the headings
            next if col_pos == 0
            attr = attribute_sym head
            value = value_from_cell cell
            next if attr.nil?
            next if value.nil?
            # each column represents a record, insert its value in the @data
            # array at the column position
            (@data[col_pos-1] ||= {})[attr] = value
          end
        end
      else
        worksheet.sheet_data.rows.each do |row|
          # don't process the first row; it has the headings
          next if row.index_in_collection == 0
          row_hash = {}
          row.cells.each_with_index do |cell, row_pos|
            attr = attribute_sym headers[row_pos]
            value = value_from_cell cell
            next if attr.nil?
            next if value.nil?
            row_hash[attr] = value
          end
          @data << row_hash
        end
      end

      @data
    end

    def attributes
      return @attributes unless @attributes.nil?

      (@sheet_config[:attributes] || []).map { |a| Attr.new deets: a }
    end

    def attribute_sym head
      return if head.nil?
      return head.to_sym unless header_map[head]
      header_map[head].attr_sym
    end

    def value_from_cell cell
      return if cell.nil?
      return if cell.value.nil?
      cell.value.to_s
    end

    def header_map
      return @header_map unless @header_map.nil?

      @header_map = attributes.inject({}) { |memo, attr|
        attr.headings.each { |h| memo[h] = attr }
        memo
      }
    end

    def required_attributes
      return @required_attributes unless @required_attributes.nil?

      @required_attributes = attributes.select &:required?
    end

    ##
    # Return the headers values for the first row or column and their positions.
    # Where a header is blank or `nil`, `nil` is in the array position.
    # For example, if there is blank header value between 'ILL ENTRY' and
    # 'FILENAME', the following might be returned.
    #
    #     [
    #       'ARK ID',
    #       'PAGE SEQUENCE',
    #       'VISIBLE PAGE',
    #       'TOC ENTRY',
    #       'ILL ENTRY',
    #       nil,
    #       'FILENAME',
    #       'NOTES'
    #     ]
    #
    # @return [Array]
    def headers
      return @headers unless @headers.empty?

      xlsx = RubyXL::Parser.parse xlsx_path
      worksheet = xlsx[@sheet_config[:sheet_position] || 0]

      if @sheet_config.fetch(:heading_type, :row).to_sym == :column
        @headers = worksheet.sheet_data.rows.map do |row|
          next nil if row.nil?
          # headers are in the first column; get the first cell value in each
          # row
          header_from_cell row.cells.first
        end
      else
        @headers = worksheet.sheet_data.rows.first.cells.map do |cell|
          header_from_cell cell
        end
      end
    end

    ##
    # Make sure there are no duplicate headers and that all the required
    # headers are present.
    # @raise [StandardError] if there are non-unique header names
    # @raise [StandardError] if one or more required columns is missing
    def validate_headers
      compact_headers = headers.compact # remove nils
      unless compact_headers.length == compact_headers.uniq.length
        raise StandardError, "Duplicate column names in #{compact_headers.sort} (#{xlsx_path})"
      end

      # # get a list of expected headers not found in `headers`
      # missing = required_headers.reject {|h| headers.include? h}
      # unless missing.empty?
      #   raise StandardError, "Missing required columns: #{missing.join '; '}"
      # end

      missing = required_attributes.reject { |a|
        a.headings.any? { |header| headers.include? header }
      }

      unless missing.empty?
        raise StandardError, "Missing required headings: #{missing.map(&:to_s).join '; '}"
      end
    end

    protected


    ##
    # @param [RubyXL::Cell] cell cell to extract the header name from
    # @return [String] the header value or nil if cell empty
    def header_from_cell cell
      return if cell.nil?
      return if cell.value.nil?
      return if cell.value.to_s.strip.empty?
      cell.value.to_s.upcase.strip
    end

    class Attr
      attr_accessor :attr, :headings, :requirement

      def initialize deets:
        @attr        = deets[:attr]
        @headings    = deets[:headings]
        @requirement = deets[:requirement]
      end

      def required?
        return unless @requirement
        return unless @requirement.is_a? String
        requirement.strip.downcase == 'required'
      end

      def to_s
        "#{attr}: (#{headings.join ', '})"
      end

      def attr_sym
        attr.to_sym
      end
    end
  end
end
