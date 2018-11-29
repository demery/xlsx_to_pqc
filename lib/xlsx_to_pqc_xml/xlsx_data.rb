require 'rubyXL'
require 'set'

module XlsxToPqcXml
  ##
  # Class to validate and extract data from a spreadsheet based on the
  # rules and configuration provided in hash.
  #
  # == Sample usage
  #
  # Take the sample configuration YAML file and sample spreadsheet below.
  #
  # Spreadsheet configuration +structural_config.yml+:
  #
  #     ---
  #     :sheet_name: 'Structural'
  #     :sheet_position: 0
  #     :heading_type: 'row'
  #     :attributes:
  #     - :attr: ark_id
  #       :headings:
  #       - ARK ID
  #       :requirement: required
  #       :data_type: :ark
  #     - :attr: page_sequence
  #       :headings:
  #       - PAGE SEQUENCE
  #       :requirement: required
  #       :unique: true
  #       :data_type: :integer
  #     - :attr: filename
  #       :headings:
  #       - FILENAME
  #       :requirement: required
  #     - :attr: visible_page
  #       :headings:
  #       - VISIBLE PAGE
  #       :requirement: required
  #     - :attr: toc_entry
  #       :headings:
  #       - TOC ENTRY
  #       :multivalued: true
  #       :value_sep: '|'
  #     - :attr: ill_entry
  #       :headings:
  #       - ILL ENTRY
  #       :multivalued: true
  #       :value_sep: '|'
  #     - :attr: notes
  #       :headings:
  #       - NOTES
  #
  # Spreadsheet './ark\+=99999=fk42244n9f/pqc_structural.xlsx':
  #
  #     | ARK ID                | PAGE SEQUENCE | VISIBLE PAGE | TOC ENTRY                | ILL ENTRY                                               | FILENAME | NOTES |
  #     |-----------------------|---------------|--------------|--------------------------|---------------------------------------------------------|----------|-------|
  #     | ark:/99999/fk42244n9f | 1             | 1r           | Pio, Alberto (1512-1518) |                                                         | 0001.tif |       |
  #     | ark:/99999/fk42244n9f | 2             | 1v           |                          |                                                         | 0002.tif |       |
  #     | ark:/99999/fk42244n9f | 3             | 2r           |                          |                                                         | 0003.tif |       |
  #     | ark:/99999/fk42244n9f | 4             | 2v           | Table, f. 2v [=3v]       |                                                         | 0004.tif |       |
  #     | ark:/99999/fk42244n9f | 5             | 3r           |                          |                                                         | 0005.tif |       |
  #     | ark:/99999/fk42244n9f | 6             | 3v-4r        |                          | Decorated initial, Initial P, p. 3|Foliate design, p. 3 | 0006.tif |       |
  #     | ark:/99999/fk42244n9f | 7             | 4v           |                          |                                                         | 0007.tif |       |
  #
  #
  # Validate and extract the data as an array of hashes:
  #
  #     require 'xlsx_to_pqc_xml'
  #     require 'yaml'
  #
  #     config = YAML.load open('structural_config.yml').read
  #     xlsx_data = XlsxToPqcXml::XlsxData.new xlsx_data: './ark+=99999=fk42244n9f/pqc_structural.xlsx', config: config
  #     xlsx_data.data
  #     # => [
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"1",
  #     #      :visible_page=>"1r",
  #     #      :toc_entry=>["Pio, Alberto (1512-1518)"],
  #     #      :filename=>"0001.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"2",
  #     #      :visible_page=>"1v",
  #     #      :filename=>"0002.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"3",
  #     #      :visible_page=>"2r",
  #     #      :filename=>"0003.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"4",
  #     #      :visible_page=>"2v",
  #     #      :toc_entry=>["Table, f. 2v [=3v]"],
  #     #      :filename=>"0004.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"5",
  #     #      :visible_page=>"3r",
  #     #      :filename=>"0005.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"6",
  #     #      :visible_page=>"3v-4r",
  #     #      :ill_entry=>["Decorated initial, Initial P, p. 3", "Foliate design, p. 3"],
  #     #      :filename=>"0006.tif"},
  #     #     {:ark_id=>"ark:/99999/fk42244n9f",
  #     #      :page_sequence=>"7",
  #     #      :visible_page=>"4v",
  #     #      :filename=>"0007.tif"}
  #     #     ]
  #
  #     xlsx_data.errors # => {} empty errors hash
  #
  # You can do validation alone by:
  #
  #     config = YAML.load open('structural_config.yml').read
  #     xlsx_data = XlsxToPqcXml::XlsxData.new xlsx_data: './ark+=99999=fk42244n9f/pqc_structural.xlsx', config: config
  #     xlsx_data.valid? # => true
  #     xlsx_data.errors # => {}
  #
  # You can extract data from a spreadsheet without validation by:
  #
  #     config = YAML.load open('structural_config.yml').read
  #     xlsx_data = XlsxToPqcXml::XlsxData.new xlsx_data: './ark+=99999=fk42244n9f/pqc_structural.xlsx', config: config
  #     xlsx_data.data data_only: true # => [{:ark_id=>...}]
  #
  # === Validation
  #
  # Validation does not halt processing unless headers fail validation. All
  # validation errors are added to an errors hash, keyed by the error type.
  #
  # For example:
  #
  #   {
  #     :non_valid_ark=>
  #       [#<OpenStruct address="A2", text="'arkx:/99999/fk42244n9f'">],
  #     :required_value_missing=>
  #       [#<OpenStruct address="C5", text="visible_page (VISIBLE PAGE)">,
  #        #<OpenStruct address="C6", text="visible_page (VISIBLE PAGE)">],
  #     :non_unique_value=>
  #       [#<OpenStruct address="B6", text="'4'; heading: page_sequence (PAGE SEQUENCE)">],
  #     :non_valid_integer=>
  #       [#<OpenStruct address="B7", text="'6.3'">]
  #    }
  #
  # Under each key is an array of OpenStruct objects with attributes +#address+
  # and +#text+. The +#address+ attribute is not present for errors that cannot
  # have an address, like +:required_header_missing+.
  #
  # Error types are:
  #
  #   :required_header_missing
  #   :non_unique_header
  #
  #   :required_value_missing
  #   :non_unique_value
  #   :non_valid_integer
  #   :non_valid_ark
  #   :non_valid_url
  #   :non_valid_web_url
  #
  # Spreadsheet headers are validated first and processing stops if the headers
  # are not valid. There are two types of header requirements:
  #
  # 1. headers must be unique
  # 2. headers for required fields must be present
  #
  # Cell values are validated according the rules described below under
  # configuration. If headers pass va
  #
  # == Configuration
  #
  # Configuration is passed as hash, but YAML will be used hear for ease of
  # reading.
  #
  #     ---
  #     :sheet_name: 'Structural'
  #     :sheet_position: 0
  #     :heading_type: 'row'
  #     :attributes:
  #     - :attr: ark_id
  #       :headings:
  #       - ARK ID
  #       :requirement: required
  #       :data_type: :ark
  #     - :attr: page_sequence
  #       :headings:
  #       - PAGE SEQUENCE
  #       :requirement: required
  #       :unique: true
  #       :data_type: :integer
  #     - :attr: filename
  #       :headings:
  #       - FILENAME
  #       :requirement: required
  #     - :attr: visible_page
  #       :headings:
  #       - VISIBLE PAGE
  #       :requirement: required
  #     - :attr: toc_entry
  #       :headings:
  #       - TOC ENTRY
  #       :multivalued: true
  #       :value_sep: '|'
  #     - :attr: ill_entry
  #       :headings:
  #       - ILL ENTRY
  #       :multivalued: true
  #       :value_sep: '|'
  #     - :attr: notes
  #       :headings:
  #       - NOTES
  #
  # ==== Top-level configuration values
  #
  # +:sheet_name+::     Any string to name the sheet. *Is not to find the
  #                     sheet.*
  #
  # +:sheet_position+:: Integer index of the worksheet in the workbook; +0+
  #                     by default
  #
  # +:heading_type+::   'row' (the default) if the headings are in the first row
  #                     of the sheet; 'column' if the headings are the first
  #                     column of the sheet
  #
  # +:attributes+::     a list of attributes to extract from the spreadsheet
  #
  # ==== Attribute configuration values
  #
  # +:attr+::           The key to store the record value under in the returned
  #                     hash
  #
  # +:headings+::       Array of possibles names for the heading; should be
  #                     unique
  #
  # +:requirement+::    Rules for whether the value is required; at present
  #                     only +required+ is evaluated, but it's intended that
  #                     more complex rules for conditional requirements could
  #                     be added
  #
  # +:unique+::         whether value must be unique under its heading
  #
  # +:multivalued+::    whether to return the contents of the cell as an array
  #
  # +:value_sep+::      separator used to split multivalued cells; +|+ (the
  #                     pipe character by default)
  #
  # Note: Columns/rows without headings are ignored. All other records are
  # extracted. Headings not listed in the attributes will be extracted using
  # their heading name; e.g., 'EXTRA HEADING' becomes +:'EXTRA HEADING'+.
  #
  # ==== Custom configuration values
  #
  # <b>XlsxData ignores custom configuration values</b>, but a calling class can
  # take advantage of them for its own purposes. For example, if you want to
  # map several attributes to XML elements, you might do something like the
  # following:
  #
  #     ---
  #     :sheet_name: 'Descriptive'
  #     :sheet_position: 0
  #     :heading_type: 'row'
  #     :attributes:
  #     - :attr: object_type
  #       :xml_element: type
  #       :headings:
  #       - OBJECT TYPE
  #       :requirement: required
  #     - :attr: unique_identifier
  #       :xml_element: ark
  #       :data_type: :ark
  #       :headings:
  #       - UNIQUE IDENTIFIER
  #       :requirement: required
  #     - :attr: abstract
  #       :xml_element: abstract
  #       :headings:
  #       - ABSTRACT
  #     - :attr: call_number
  #       :xml_element: callNumber
  #       :headings:
  #       - CALL NUMBER
  #     # ... etc.
  #
  # Here, +:attr+ 'object_type' has +:xml_element+ name 'type'.
  #
  #
  class XlsxData

    attr_reader :xlsx_path
    attr_reader :errors

    ##
    # Create a new XlsxData for XLSX file +xlsx_path+ with +config+ hash.
    #
    # Config hash is similar to the following:
    #
    #     {:sheet_name=>"Structural",
    #      :sheet_position=>0,
    #      :heading_type=>"row",
    #      :attributes=>
    #        [{
    #           :attr=>"ark_id",
    #           :headings=>["ARK ID"],
    #           :requirement=>"required",
    #           :data_type=>:ark
    #         },
    #         {
    #           :attr=>"page_sequence",
    #           :headings=>["PAGE SEQUENCE"],
    #           :requirement=>"required",
    #           :unique=>true,
    #           :data_type=>:integer
    #         },
    #         {
    #           :attr=>"filename",
    #           :headings=>["FILENAME"],
    #           :requirement=>"required"
    #         },
    #         {
    #           :attr=>"visible_page",
    #           :headings=>["VISIBLE PAGE"],
    #           :requirement=>"required"
    #         },
    #         {
    #           :attr=>"toc_entry",
    #           :headings=>["TOC ENTRY"],
    #           :multivalued=>true,
    #           :value_sep=>"|"
    #         },
    #         {
    #           :attr=>"ill_entry",
    #           :headings=>["ILL ENTRY"],
    #           :multivalued=>true,
    #           :value_sep=>"|"
    #         },
    #         {
    #           :attr=>"notes",
    #           :headings=>["NOTES"]
    #         }
    #       ]
    #     }
    #
    # @param [String] xlsx_path path to the XLSX file
    # @param [Hash] config spreadsheet configuration
    def initialize xlsx_path:, config: {}
      @xlsx_path           = xlsx_path
      @sheet_config        = config.dup # be defensive
      @data                = []
      @header_addresses    = []
      @errors              = Hash.new { |hash, key| hash[key] = [] }
      @required_attributes = []
      @required_headings   = []
      @attributes          = []
      @header_map          = {}
      @extracted           = false
    end

    ##
    # Process the spreadsheet without extracting data and validate
    #
    # @return [Boolean] true if there are no errors
    def valid?
      @errors.clear

      unless validate_config
        raise XlsxDataException.new "Invalid configuration", errors = @errors
      end

      return unless validate_headers

      process validation_only: true
      @errors.empty?
    end

    ##
    # Extract data and validate spreadsheet.
    #
    # @param [Boolean] data_only skip validation
    # @return [Array<Hash>] array of the spreadsheet data as hashes
    def data data_only: false
      return @data unless @data.empty?
      process data_only: data_only
      @data
    end

    ##
    # Read and validate the spreadsheet at {xlsx_path}.
    #
    # NOTE: Data will be an empty array if +:validation_only+ is +true+.
    #
    # @param [Boolean] data_only skip validation
    # @param [Boolean] validation_only skip data extraction
    # @return [Array<Hash>] array of the spreadsheet data as hashes
    def process data_only: false, validation_only: false

      reset # clear all cached data

      @data = []

      xlsx      = RubyXL::Parser.parse xlsx_path
      worksheet = xlsx[@sheet_config[:sheet_position] || 0]
      uniques   = Hash.new { |hash, key| hash[key] = Set.new }

      unless data_only
        validate_headers
        unless @errors.empty?
          $stderr.puts "WARNING: Invalid headers! Processing aborted." unless @errors.empty?
          return unless @errors.empty?
        end
      end

      if @sheet_config.fetch(:heading_type, :row).to_sym == :column
        # headings are in the first column; for each header, work across the
        # row, collecting the value in each column.
        headers.each_with_index do |head, row_pos|
          next if head.nil?
          worksheet.sheet_data.rows[row_pos].cells.each_with_index do |cell, col_pos|
            next if col_pos == 0 # skip header column

            # each column represents a record, insert its value in the @data
            # array at the column position
            row_hash = @data[col_pos - 1] ||= {}

            cell_data                 = CellParams.new
            cell_data.cell            = cell
            cell_data.row_pos         = row_pos
            cell_data.col_pos         = col_pos
            cell_data.head            = head
            cell_data.data_only       = data_only
            cell_data.validation_only = validation_only
            cell_data.row_hash        = row_hash

            process_cell cell_data, uniques
          end
        end
      else
        worksheet.sheet_data.rows.each_with_index do |row, row_pos|
          next if row_pos == 0 # skip header row

          row_hash = {}
          row.cells.each_with_index do |cell, col_pos|
            cell_data                 = CellParams.new
            cell_data.cell            = cell
            cell_data.row_pos         = row_pos
            cell_data.col_pos         = col_pos
            cell_data.head            = headers[col_pos]
            cell_data.data_only       = data_only
            cell_data.validation_only = validation_only
            cell_data.row_hash        = row_hash

            process_cell cell_data, uniques
          end
          @data << row_hash unless validation_only
        end
      end
      @extracted = true unless validation_only

      @data
    end

    ###########################################################################
    # CONFIGURATION
    ###########################################################################

    ##
    # Return an array of {Attr} instances for the given sheet config.
    #
    # @return [Array<Attr>] all configured attributes
    def attributes
      return @attributes unless @attributes.empty?

      (@sheet_config[:attributes] || []).map { |a| Attr.new deets: a }
    end

    ##
    # Return an array of {Attr} instances for all configured attributes where
    # the {Attr#required?} returns +true+; for example, if:
    #
    #   @sheet_config[:attributes][0][:requirement] == 'required'
    #
    # then that attribute configuration would be returned in the array.
    #
    # @return [Array<Attr>] all required attributes
    def required_attributes
      return @required_attributes unless @required_attributes.empty?

      @required_attributes = attributes.select &:required?
    end

    def required_headings
      return @required_headings unless @required_headings.empty?

      @required_headings = attributes.select &:heading_required
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
    def valid_header_map
      return @header_map unless @header_map.empty?

      @header_map = attributes.inject({}) { |memo, attr|
        # map under all case possibilities
        attr.headings.each          { |h| memo[h] = attr }
        attr.upcase_headings.each   { |h| memo[h] = attr }
        attr.downcase_headings.each { |h| memo[h] = attr }
        memo
      }
    end

    def find_attr head
      return if head.nil?
      valid_header_map[head.to_s.upcase]
    end

    ###########################################################################
    # CELL HANDLING
    ###########################################################################

    ##
    # Extract and/or validate the given cell.
    #
    # @param [CellParams] cell_data all values needed to process a cell
    # @param [Hash] uniques hash to track unique_values
    def process_cell cell_data, uniques
      attr = find_attr cell_data.head
      attr_sym = attribute_sym cell_data.head
      return if attr_sym.nil?

      unless cell_data.data_only
        address = cell_address(cell_data.col_pos, cell_data.row_pos)
        return unless cell_valid? cell_data.cell, attr, address, uniques
      end

      return if cell_data.validation_only
      value = value_from_cell cell_data.cell, cell_data.head
      return if value.nil?

      cell_data.row_hash[attr_sym] = value
    end

    ##
    # Return the value of the cell, splitting the cell if it is multi-valued.
    # Returns +nil+ if +head+ nil or +cell+ is empty.
    #
    # @param [RubyXL::Cell]  cell the cell containing the data
    # @param [String] head heading value for the cell's column/row
    def value_from_cell cell, head
      return if head.nil?
      attr = find_attr head

      val = bare_cell_value cell
      return if val.nil?
      return val unless attr.is_a? Attr
      return val unless attr.multivalued?

      val.split(attr.split_regex).map(&:strip)
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
      headers_with_addresses.map &:header
    end

    ##
    # Return all the spreadsheet headers with their addresses in Excel format;
    # e.g., 'A1', 'A2', etc. Return value is an array of OpenStruct instances
    # with attributes `#header` and `#address`.
    #
    # @return [Array<OpenStruct>] array of header value and addresses.
    def headers_with_addresses
      return @header_addresses unless @header_addresses.empty?
      @header_addresses = []

      xlsx = RubyXL::Parser.parse xlsx_path
      worksheet = xlsx[@sheet_config[:sheet_position] || 0]

      if @sheet_config.fetch(:heading_type, :row).to_sym == :column
        col_pos = 0
        worksheet.sheet_data.rows.each_with_index do |row, row_pos|
          # headers are in the first column; get the first cell value in each row
          cell = row.cells.first unless row.nil?
          address = cell_address col_pos, row_pos

          val = header_from_cell cell
          @header_addresses << OpenStruct.new(header: val, address: address)
        end
      else
        row_pos = 0
        worksheet.sheet_data.rows.first.cells.each_with_index do |cell, col_pos|
          address = cell_address col_pos, row_pos
          val = header_from_cell cell

          @header_addresses << OpenStruct.new(header: val, address: address)
        end
      end
      @header_addresses
    end

    ##
    # @return [Boolean] true if data has been extracted
    def extracted?
      @extracted
    end

    ##
    # Reset instance to pre-processed, validation state, clearing all cached
    # instance variables
    def reset
      @data.clear
      @header_addresses.clear
      @errors.clear
      @required_attributes.clear
      @required_headings.clear
      @attributes.clear
      @header_map.clear
      @extracted = false
    end

    ###########################################################################
    # VALIDATION
    ###########################################################################

    # Class methods
    class << self

      # regular expression to match an ark
      ARK_REGEX = %r{\Aark:/\w+/\w+\Z}i

      # Hash of data type validators
      DEFAULT_TYPE_VALIDATORS = {
        integer: lambda { |value|
          begin
            Integer value
          rescue ArgumentError
            false
          end
        },
        ark:     lambda { |value| value =~ ARK_REGEX },
        uri:     lambda { |value| value =~ URI.regexp },
        web_url: lambda { |value| value =~ URI.regexp(%w{http https}) },
      }.freeze

      @@type_validators = DEFAULT_TYPE_VALIDATORS.inject({}) { |memo, type_lambda|
        memo[type_lambda.first] = type_lambda.last
        memo
      }

      ##
      # @return [Hash] copy of currently configure validators
      def type_validators
        @@type_validators.dup
      end

      ##
      # Add a new or replace an existing +validator+ for +data_type+. The
      # validator must be a lambda that takes one argument, the cell value
      # and returns a truthy value if value is valid and non-truthy otherwise.
      #
      # The validator should not be passed a +nil+ value and so does not need to
      # handle the case of +nil+.
      #
      # @param [Symbol] data_type e.g., +:string+, +:year+
      # @param [Lambda] validator
      def set_type_validator data_type, validator
        @@type_validators[data_type.to_sym] = validator
      end

      ##
      # Delete the validator for +data_type+.
      #
      # @param [Symbol] data_type
      # @return [Lambda]
      def delete_type_validator data_type
        @@type_validators.delete data_type.to_sym
      end

      ##
      # Return the validator for +data_type+.
      #
      # @param [Symbol] data_type
      # @return [Lambda] the deleted validator
      def type_validator data_type
        @@type_validators[data_type.to_sym]
      end

      ##
      # Return whether a validator exists for +data_type+.
      #
      # @param [Symbol] data_type
      # @return [Boolean]
      def has_validator? data_type
        @@type_validators.include? data_type.to_sym
      end
    end

    ##
    # Make sure there are no duplicate headers and that all the required
    # headers are present.
    #
    # @return [Boolean] true if all header validations pass
    def validate_headers
      valid = true
      valid &= validate_headers_unique
      valid &= validate_required_headers

      valid
    end

    ##
    # If there are non-unique headers, add the +:non_unique_header+ error to
    # errors hash for each header and return +false+; otherwise, return +true+.
    #
    # @return [Boolean] true if all headers are unique
    def validate_headers_unique
      compact_headers = headers.compact.map &:upcase # remove nils
      return true if compact_headers.length == compact_headers.uniq.length

      header_count = headers_with_addresses.inject({}) { |memo,struct|
        (memo[struct.header.upcase] ||= []) << struct unless struct.header.nil?
        memo
      }
      header_count.each do |head, structs|
        next unless structs.size > 1
        add_error :non_unique_header, "#{head}", structs.map(&:header)
      end

      false
    end

    ##
    # If there are non-unique headers, add a +:required_header_missing+ error to
    # errors hash for each header and return +false+; otherwise, return +true+.
    #
    # @return [Boolean] true if all required headers are present
    def validate_required_headers
      upcase_headers = headers.map { |h| h && h.upcase }
      missing = required_attributes.reject { |a|
        a.upcase_headings.any? { |header| upcase_headers.include? header }
      }
      # binding.pry

      missing += required_headings.reject { |a|
        a.upcase_headings.any? { |header| upcase_headers.include? header }
      }
      # binding.pry
      return true if missing.empty?

      missing.each do |head|
        msg = "Required header not found: #{head}"
        add_error :required_header_missing, msg
      end

      false
    end

    ##
    # If cell has an attribute configuration, check it for validity:
    # defined requirement, uniqueness, or data type constraints.
    #
    # Invokes the following validation methods and returns false at the first
    # failure encountered:
    #
    #     validate_requirement
    #     validate_uniqueness
    #     validate_type
    #
    # @param [RubyXL::Cell] cell
    # @param [Attr] attr
    # @param [String] address the Excel style cell address; .e.g, 'A2'
    # @param [Hash] uniques cache of values for validating uniqueness
    # @return [Boolean] false if an validation fails
    def cell_valid? cell, attr, address, uniques
      return if attr.nil?
      value = bare_cell_value cell

      return false unless validate_requirement value, attr, address
      return false unless validate_uniqueness value, attr, address, uniques
      return false unless validate_type value, attr, address

      true
    end

    ##
    # If +value+ is present and +attr#unique?+ is +true+, add the error to
    # errors hash and return +false+; otherwise, return +true+.
    #
    # @param [String] value the cell value
    # @param [Attr] attr the attribute configuration
    # @param [String] address Excel style cell address; e.g., 'A2'
    #
    # @return [Boolean] true if the value passes validation
    def validate_uniqueness value, attr, address, uniques
      return true unless attr.unique?
      return true unless value
      if uniques[attr.attr_sym].include? value
        add_error :non_unique_value, "'#{value}'; heading: #{attr}", address
        return false
      end

      uniques[attr.attr_sym] << value
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
    # @return [Boolean] true if the value passes validation
    def validate_requirement value, attr, address
      return true unless attr.required?
      return true unless value.nil?

      add_error :required_value_missing, "#{attr}", address
      false
    end

    ##
    # If +value+ and +attr#data_type+ are present, return true if +value+
    # passes validation.
    #
    # @param [String] value the cell value
    # @param [Attr] attr the attribute configuration
    # @param [String] address Excel style cell address; e.g., 'A2'
    #
    # @return [Boolean] true if the value passes validation
    # @raise [XlsxDataException] if data_type is not known
    def validate_type value, attr, address
      return true unless attr.data_type
      return true unless value
      data_type = attr.data_type
      validator = XlsxData.type_validator data_type
      raise XlsxDataException, "Unknown data type: #{data_type}" unless validator
      return true if validator.call value

      error_sym = "non_valid_#{data_type}".to_sym
      add_error error_sym, "'#{value}'", address
      false
    end

    ##
    # Make sure we have a valid configuration:
    #
    # - All configuration data types must be defined
    # - All attributes must have an :attr
    # - All attributes must have an Array of valid :headings
    #
    # @return [Boolean] false if any checks fail
    def validate_config
      valid = true
      attributes.each do |attr|
        next unless attr.data_type
        unless XlsxData.has_validator? attr.data_type
          add_error :unknown_data_type, attr.data_type
          valid &= false
        end
      end

      # TODO: Add check for duplicate attrs
      # TODO: Add check for duplicate headings
      @sheet_config.fetch(:attributes, []).each do |attr_hash|

        unless attr_hash[:attr]
          add_error :attr_not_defined, attr_hash.inspect
          valid &= false
        end

        unless attr_hash[:headings].is_a? Array
          add_error :no_headings_array, attr_hash.inspect
          valid &= false
        end

      end

      valid
    end

    ##
    # Add a single error to the errors array stored under the key +error_sym+.
    #
    # Each error is added to an array in the errors hash under the key
    # +error_sym+. The error information is stored in an OpenStruct instance
    # with attributes +#address+ and +#text+. A typical key value pair would look
    # like:
    #
    #     :non_valid_integer  => [#<OpenStruct address="B3", text="'2.3'">]
    #
    # @param [Symbol] error_sym a key identify the error type, e.g., +:required_value_missing+
    # @param [String,Array<String>] address Excel style address (e.g.,
    #                                     'A3') or array of addresses
    # @param [String] text clarifying information like the name of the missing header
    # @return [Array<OpenStruct>] current array of error structs under key +error_sym+
    def add_error error_sym, text, address=nil
      @errors[error_sym] << OpenStruct.new(address: address, text: text)
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
    }

    def cell_address col_index, row_index
      "#{COLUMN_INDEX_TO_LETTER[col_index]}#{row_index + 1}"
    end


    ###########################################################################
    # HELPER CLASSES
    ###########################################################################

    ##
    # Convenience class to pass cell information to {#process_cell}.
    class CellParams
      attr_accessor :cell, :row_pos, :col_pos, :row_hash, :head, :data_only, :validation_only
    end

    ##
    # Convenience class to encapsulate the configuration of an attribute,
    # with boolean convenience methods for required and multivalued fields,
    # return the attr name as a Symbol.
    #
    class Attr
      attr_accessor :attr, :headings, :requirement, :multivalued, :value_sep
      attr_accessor :unique, :data_type, :heading_required

      DEFAULT_VALUE_SEP = '|'

      def initialize deets:
        @attr             = deets[:attr]
        @headings         = deets[:headings] || []
        @requirement      = deets[:requirement]
        @multivalued      = deets[:multivalued]
        @value_sep        = deets[:value_sep] || DEFAULT_VALUE_SEP
        @unique           = deets[:unique] || false
        @data_type        = deets[:data_type] and deets[:data_type].to_s.to_sym
        @heading_required = deets[:heading_required] || false
      end

      ##
      # Return +true+ if +:requirement+ is defined and is 'required'.
      #
      # @return [Boolean]
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

      def heading_required?
        @heading_required
      end

      def upcase_headings
        headings.map &:upcase
      end

      def downcase_headings
        headings.map &:downcase
      end

      ##
      # Return the separator, if defined ,as a regular expression for splitting
      # a column; return nil otherwise.
      #
      # @return [Regexp]
      def split_regex
        return if value_sep.nil?
        return if value_sep.to_s.strip.empty?

        /#{Regexp.escape(value_sep)}/
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
      cell.value.to_s.strip
    end

    ##
    # For the given +head+ value, like 'FILENAME', return the corresponding
    # configured attribute as a symbol or the name converted to a symbol. Return
    # +nil+ if +head+ is nil.
    #
    # For example,
    #
    #   # attribute config with attr 'filename' exists for 'FILENAME'
    #   attribute_sym 'FILENAME'              # => :filename
    #
    #   # no attribute configured
    #   attribute_sym 'unconfigured header'   # => :'unconfigured header'
    #
    # @param [String] head a head value
    # @return [Symbol]
    def attribute_sym head
      return if head.nil?
      return head.to_sym unless find_attr head
      find_attr(head).attr_sym
    end
  end

  class XlsxDataException < StandardError
    attr_reader :errors

    def initialize msg='XlsxDataException encountered', errors={}
      @errors = errors
      super msg
    end
  end
end
