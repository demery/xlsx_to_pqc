require 'rubyXL'
require 'ostruct'

module XlsxToMarmite
  class StructuralMetadata
    ##
    # x - Confirm all TIFF files in XLSX in directory
    # x - Confirm all required headers present
    # TODO: Validate required values present
    # TODL: Return PQC XML

    # --- Column headers ---
    ARK_ID        = 'ARK ID'
    PAGE_SEQUENCE = 'PAGE SEQUENCE'
    VISIBLE_PAGE  = 'VISIBLE PAGE'
    TOC_ENTRY     = 'TOC ENTRY'
    ILL_ENTRY     = 'ILL ENTRY'
    FILENAME      = 'FILENAME'
    NOTES         = 'NOTES'
    HEADER_NAMES  = [
      ARK_ID, PAGE_SEQUENCE, VISIBLE_PAGE, TOC_ENTRY, ILL_ENTRY,
      FILENAME, NOTES
    ].freeze
    # Don't require 'NOTES' column
    REQUIRED_HEADERS                 = HEADER_NAMES.reject {|h| h == NOTES}.freeze

    # --- Default values --- (in case we want to make editable later)
    DEFAULT_PQC_STRUCTURAL_XLSX_BASE = 'pqc_structural.xlsx'.freeze
    DEFAULT_DATA_FILE_GLOB           = '*.{tif,tiff}'.freeze

    # --- Other constants ---
    RECTO                            = 'recto'.freeze
    VERSO                            = 'verso'.freeze

    attr_reader :package_directory
    attr_reader :expected_headers
    attr_accessor :data_file_glob

    ##
    # @param [String] package_directory Path the to content package directory
    def initialize package_directory
      @package_directory    = package_directory
      @expected_headers     = REQUIRED_HEADERS.map &:itself
      @structural_xlsx_base = DEFAULT_PQC_STRUCTURAL_XLSX_BASE
      @data_file_glob       = DEFAULT_DATA_FILE_GLOB
      @spreadsheet_data     = []
      @data_file_list       = []
      @image_data           = []
      @spreadsheet_files    = []
    end

    ##
    # @return [String] full path to the structural metadata spreadsheet
    def xlsx_path
      File.join package_directory, @structural_xlsx_base
    end

    ##
    # Return an array of struct objects with these properties for each image:
    #
    #     #number
    #     #seq
    #     #image_defaultscale
    #     #display
    #     #side
    #     #image_id
    #     #image
    #     #visiblepage
    #     #toc
    #     #ill
    #
    # Image files on disk but not in the spreadsheet are listed added at the
    # end. Their +number+ and +seq+ values follow the spreadsheet file values
    # sequentially; e.g., if the last spreadsheet value has +number+ and +seq+
    # value +6+, the first file not listed in the spreadsheet will get +number+
    # and +seq+ value +7+.
    #
    # @return [Array<OpenStruct>] list of structs with data for all images
    def image_data
      return @image_data unless @image_data.empty?

      validate_spreadsheet_file_list

      spreadsheet_data.map do |row_hash|
        sequence = Integer row_hash[PAGE_SEQUENCE].to_s.strip
        filename = row_hash[FILENAME]
        image    = File.basename filename, File.extname(filename)
        toc      = (row_hash[TOC_ENTRY] || '').split(%r{\s*\|\s*}).map(&:strip)
        ill      = (row_hash[ILL_ENTRY] || '').split(%r{\s*\|\s*}).map(&:strip)
        data     = {
          number:             sequence,
          seq:                sequence,
          image_defaultscale: 3,
          display:            true,
          side:               (sequence.odd? ? RECTO : VERSO),
          image_id:           image,
          image:              image,
          visiblepage:        row_hash[VISIBLE_PAGE],
          toc:                toc,
          ill:                ill
        }
        @image_data << OpenStruct.new(data)
      end

      files_on_disk.reject {|f| spreadsheet_files.include? f }.each do |extra|
        sequence = @image_data.last.seq + 1
        image    = File.basename extra, File.extname(extra)
        data = {
          number:             sequence,
          seq:                sequence,
          id:                 image,
          image_defaultscale: 3,
          side:               (sequence.odd? ? RECTO : VERSO),
          image_id:           image,
          image:              image,
          visiblepage:        nil,
          display:            false,
          toc:                [],
          ill:                []
        }
        @image_data << OpenStruct.new(data)
      end

      @image_data
    end

    ##
    # @return [Array<Hash>] array of the spreadsheet data as hashes
    # @raise [StandardError] if there are duplicate headers or missing columns
    def spreadsheet_data
      return @spreadsheet_data unless @spreadsheet_data.empty?

      @spreadsheet_data = []

      xlsx      = RubyXL::Parser.parse xlsx_path
      worksheet = xlsx[0]
      headers   = worksheet.sheet_data.rows.first.cells.map do |cell|
        cell.value.upcase.strip
      end
      validate_headers headers

      worksheet.sheet_data.rows.each do |row|
        next if row.index_in_collection == 0
        row_hash = {}
        row.cells.each_with_index do |cell, position|
          value                       = (cell.nil? || cell.value.nil?) ? '' : cell.value.to_s
          row_hash[headers[position]] = value
        end
        @spreadsheet_data << row_hash
      end

      @spreadsheet_data
    end

    ##
    # @return [Array<String>] list of data file basenames in {package_directory}
    def files_on_disk
      return @data_file_list unless @data_file_list.empty?

      Dir.chdir package_directory do
        @data_file_list = Dir[data_file_glob]
      end

      @data_file_list
    end

    ##
    # @return [Array<String>] list files in spreadsheet FILENAME column
    def spreadsheet_files
      return @spreadsheet_files unless @spreadsheet_files.empty?

      @spreadsheet_files = spreadsheet_data.flat_map do |row_hash|
        next [] if row_hash[FILENAME].nil?
        next [] if row_hash[FILENAME].to_s.strip.empty?
        row_hash[FILENAME].to_s
      end
    end

    protected

    ##
    # Make sure there are no duplicate headers and that all the required
    # headers are present.
    # @raise [StandardError] if there are non-unique header names
    # @raise [StandardError] if one or more required columns is missing
    def validate_headers headers
      unless headers.length == headers.uniq.length
        raise StandardError, "Duplicate column names in #{headers.sort} (#{xlsx_path})"
      end

      # get a list of expected headers not found in `headers`
      missing = expected_headers.reject {|h| headers.include? h}
      unless missing.empty?
        raise StandardError, "Missing required columns: #{missing.join '; '}"
      end
    end

    ##
    # Make sure we have all the files listed in the spreadsheet
    # @raise [StandardError] if the spreadsheet lists files not found on disk
    def validate_spreadsheet_file_list
      missing = spreadsheet_files.reject {|file| files_on_disk.include? file}

      unless missing.empty?
        list = missing.map {|f| "'#{f}'"}.join ', '
        raise StandardError, "Spreadsheet files not found in folder: #{list}"
      end
    end

  end
end
