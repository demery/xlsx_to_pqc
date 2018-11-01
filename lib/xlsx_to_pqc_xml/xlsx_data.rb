require 'rubyXL'
require 'set'

module XlsxToPqcXml
  class XlsxData

    attr_reader :xlsx_path
    attr_reader :errors

    ARK_REGEX = %r{\Aark:/\w+/\w+\Z}i

    # TODO: Validate required values present
    # TODO: Validate pattern/data types
    # TODO: Validate values uniq

    ##
    # Create a new XlsxData for XLSX file +xlsx_path+ with +config+ hash.
    #
    # Config hash is similar to the following:
    #
    #     { :sheet_name=>"Structural",
    #       :sheet_position=>0,
    #       :heading_type=>"column",
    #       :attributes=>
    #         [{ :attr=>"ark_id",
    #            :headings=>["ARK ID"],
    #            :requirement=>"required"
    #          },
    #          { :attr=>"page_sequence",
    #            :headings=>["PAGE SEQUENCE"],
    #            :requirement=>"required"
    #          },
    #          { :attr=>"filename",
    #            :headings=>["FILENAME"],
    #            :requirement=>"required"
    #          },
    #          { :attr=>"visible_page",
    #            :headings=>["VISIBLE PAGE"],
    #            :requirement=>"required"
    #          },
    #          { :attr=>"toc_entry",
    #            :headings=>["TOC ENTRY"],
    #            :multivalued=>true,
    #            :value_sep=>"|"
    #          },
    #          { :attr=>"ill_entry",
    #            :headings=>["ILL ENTRY"],
    #            :multivalued=>true,
    #            :value_sep=>"|"
    #          },
    #          { :attr=>"notes",
    #            :headings=>["NOTES"]
    #          }
    #         ]
    #      }
    #
    # @param [String] xlsx_path path to the XLSX file
    # @param [Hash] config spreadsheet configuration
    def initialize xlsx_path:, config: {}
      @xlsx_path        = xlsx_path
      @sheet_config     = config
      @data             = nil
      @headers          = []
      @errors           = Hash.new { |hash, key| hash[key] = [] }
      @uniques          = Hash.new { |hash, key| hash[key] = Set.new }
    end

    def valid?
      data
      @errors.empty?
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
          worksheet.sheet_data.rows[row_pos].cells.each_with_index do |cell, col_pos|
            # don't process the first column; it has the headings
            next if col_pos == 0
            # each column represents a record, insert its value in the @data
            # array at the column position
            row_hash = @data[col_pos - 1] ||= {}

            attr = header_map[head]
            attr_sym = attribute_sym head
            next unless cell_valid? cell, attr, cell_address(col_pos, row_pos)
            value = value_from_cell cell, head
            next if attr_sym.nil?
            next if value.nil?
            row_hash[attr_sym] = value
          end
        end
      else
        worksheet.sheet_data.rows.each do |row|
          # don't process the first row; it has the headings
          row_pos = row.index_in_collection
          next if row_pos == 0
          row_hash = {}
          row.cells.each_with_index do |cell, col_pos|

            attr = header_map[headers[col_pos]]
            attr_sym = attribute_sym headers[col_pos]
            next unless cell_valid? cell, attr, cell_address(col_pos, row_pos)
            value = value_from_cell cell, headers[col_pos]
            next if attr_sym.nil?
            next if value.nil?
            row_hash[attr_sym] = value
          end
          @data << row_hash
        end
      end

      @data
    end

    ##
    # Return an array of {Attr} instances for the given sheet config.
    #
    # @return [Array<Attr>] all configured attributes
    def attributes
      return @attributes unless @attributes.nil?

      (@sheet_config[:attributes] || []).map { |a| Attr.new deets: a }
    end

    ##
    # Return and array of {Attr} instances for all configured attributes where
    # the {Attr#required?} returns +true+; for example, if:
    #
    #   @sheet_config[:attributes][0][:requirement] == 'required'
    #
    # then that attribute configuration would be returned in the array.
    #
    # @return [Array<Attr>] all required attributes
    def required_attributes
      return @required_attributes unless @required_attributes.nil?

      @required_attributes = attributes.select &:required?
    end

    ##
    # Return the value of the cell, splitting the cell if it is multi-valued.
    # Returns +nil+ if +head+ nil or +cell+ is empty.
    #
    # @param [RubyXL::Cell]  cell the cell containing the data
    # @param [String] head heading value for the cell's column/row
    def value_from_cell cell, head
      return if head.nil?
      attr = header_map[head]

      val = bare_cell_value cell
      return if val.nil?
      return val unless attr.is_a? Attr
      return val unless attr.multivalued?

      val.split(/#{Regexp.escape(attr.value_sep)}/).map(&:strip)
    end

    def cell_valid? cell, attr, address
      return if attr.nil?
      value = bare_cell_value cell

      return false unless validate_requirement value, attr, address
      return false unless validate_uniqueness value, attr, address

      true
    end

    ##
    # If +value+ is not present and +attr#required?+ is +true+, add the
    # error to the errors hash and return +false+; otherwise, return +true+.
    #
    # @param [String] value the cell value
    # @param [Attr] attr the attribute configuration
    # @param [String] address Excel style cell address; e.g., 'A2'
    #
    # @return [Boolean] false if the value fails validation
    def validate_requirement value, attr, address
      return true unless attr.required?
      return true unless value.nil?

      @errors[:required_value_missing] << [address, "#{attr}"]
      false
    end

    ##
    # If +value+ is present and +attr#unique?+ is +true+, add the error to
    # errors hash and return +false+; otherwise, return +true+.
    #
    # @param [String] value the cell value
    # @param [Attr] attr the attribute configuration
    # @param [String] address Excel style cell address; e.g., 'A2'
    #
    # @return [Boolean] false if the value fails validation
    def validate_uniqueness value, attr, address
      return true unless attr.unique?
      return true unless value
      if @uniques[attr.attr_sym].include? value
        @errors[:non_unique_value] << [address, "'#{value}'; heading: #{attr}"]
        return false
      end

      @uniques[attr.attr_sym] << value
      true
    end
    ##
    # Return the cell value as a string; return nil if cell is blank.
    #
    # @param [RubyXL::Cell] cell
    # @return [String]
    def bare_cell_value cell
      return if cell.nil?
      return if cell.value.nil?
      return if cell.value.to_s.strip.empty?

      cell.value.to_s.strip
    end

    ##
    # Return Hash of headers mapped to their {Attr} instances. For example,
    # given:
    #
    #     attr1.headers = ['Title']
    #     attr2.headers = ['Filename', 'File name']
    #
    # Return the following hash:
    #
    #     {
    #       'Title'       => attr1,
    #       'Filename'    => attr2,
    #       'File name',  => attr2
    #     }
    #
    # @return [Hash] mapping of each allowed header to its attribute
    def header_map
      return @header_map unless @header_map.nil?

      @header_map = attributes.inject({}) { |memo, attr|
        attr.headings.each { |h| memo[h] = attr }
        memo
      }
    end

    ##
    # Return the headers values for the first row or column and their positions.
    # Where a header is blank, +nil+ is in the array position.
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

      missing = required_attributes.reject { |a|
        a.headings.any? { |header| headers.include? header }
      }

      unless missing.empty?
        raise StandardError, "Missing required headings: #{missing.map(&:to_s).join '; '}"
      end
    end

    protected

    ##
    # Recursive Hash that returns the Excel column letter for given index:
    #
    #   0 => 'A'
    #   1 => 'B'
    #   ...
    #   26 => 'AA'
    #   etc.
    #
    COLUMN_INDEX_TO_LETTER = Hash.new { |hash, key|
      ndx = key ? key.to_i : 0
      hash[ndx] = (ndx == 0 ?  'A' : hash[ndx - 1].succ)
      hash
    }

    def cell_address col_index, row_index
      "#{COLUMN_INDEX_TO_LETTER[col_index]}#{row_index}"
    end

    ##
    # Convenience class to encapsulate the configuration of an attribute,
    # with boolean convenience methods for required and multivalued fields,
    # return the attr name as a {Symbol}.
    #
    class Attr

      # TODO: Add split RegEx method

      attr_accessor :attr, :headings, :requirement, :multivalued, :value_sep

      DEFAULT_VALUE_SEP = '|'

      def initialize deets:
        @attr        = deets[:attr]
        @headings    = deets[:headings]
        @requirement = deets[:requirement]
        @multivalued = deets[:multivalued]
        @value_sep   = deets[:value_sep] || DEFAULT_VALUE_SEP
        @unique      = deets[:unique] || false
      end

      def required?
        return unless @requirement
        return unless @requirement.is_a? String
        requirement.strip.downcase == 'required'
      end

      def multivalued?
        @multivalued
      end

      def unique?
        @unique
      end

      def to_s
        "#{attr} (#{headings.join ', '})"
      end

      def attr_sym
        attr.to_sym
      end
    end

    ##
    # @param [RubyXL::Cell] cell cell to extract the header name from
    # @return [String] the header value or nil if cell empty
    def header_from_cell cell
      return if cell.nil?
      return if cell.value.nil?
      return if cell.value.to_s.strip.empty?
      cell.value.to_s.upcase.strip
    end

    ##
    # For the given {head} value, like 'FILENAME', return the corresponding
    # configured attribute as a symbol or the name converted to a symbol. Return
    # {nil} if {head} is nil.
    #
    # For example,
    #
    #   head = 'FILENAME' # a configured header
    #   attribute_sym head # => # :filename
    #
    #   head = 'unconfigured header'
    #   attribute_sym head # => :'unconfigured header'
    #
    # @param [String] head a head value
    # @return [Symbol]
    def attribute_sym head
      return if head.nil?
      return head.to_sym unless header_map[head]
      header_map[head].attr_sym
    end
  end
end
