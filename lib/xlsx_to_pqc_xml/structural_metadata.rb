require 'ostruct'
require 'nokogiri'

module XlsxToPqcXml
  class StructuralMetadata

    # --- Default values --- (in case we want to make editable later)
    DEFAULT_PQC_STRUCTURAL_XLSX_BASE = 'pqc_structural.xlsx'.freeze
    DEFAULT_DATA_FILE_GLOB           = '*.{tif,tiff}'.freeze

    # --- Other constants ---
    RECTO                            = 'recto'.freeze
    VERSO                            = 'verso'.freeze

    attr_reader :package_directory
    attr_accessor :data_file_glob

    ##
    # @param [String] package_directory Path the to content package directory
    def initialize package_directory:, sheet_config:
      @package_directory    = package_directory
      @xlsx_data            = nil
      @structural_xlsx_base = DEFAULT_PQC_STRUCTURAL_XLSX_BASE
      @data_file_glob       = DEFAULT_DATA_FILE_GLOB
      @sheet_config         = sheet_config
      @data_file_list       = []
      @image_data           = []
      @spreadsheet_files    = []
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

      # Create structs for all the spreadsheet files
      spreadsheet_data.each do |row_hash|
        sequence = Integer row_hash[:page_sequence].to_s.strip
        filename = row_hash[:filename]
        image    = File.basename filename, File.extname(filename)
        toc      = row_hash[:toc_entry] || []
        ill      = row_hash[:ill_entry] || []
        # NB: #display is a method of Object; hence :display? below
        data     = {
          number:             sequence,
          seq:                sequence,
          image_defaultscale: 3,
          display?:           true,
          side:               (sequence.odd? ? RECTO : VERSO),
          image_id:           image,
          image:              image,
          visiblepage:        row_hash[:visible_page],
          toc:                toc,
          ill:                ill
        }
        @image_data << OpenStruct.new(data)
      end

      # Create structs for all the files not in the spreadsheet
      files_on_disk.reject {|f| spreadsheet_files.include? f }.each do |extra|
        sequence = @image_data.last.seq + 1
        image    = File.basename extra, File.extname(extra)
        # NB: #display is a method of Object; hence :display? below
        data = {
          number:             sequence,
          seq:                sequence,
          id:                 image,
          image_defaultscale: 3,
          side:               (sequence.odd? ? RECTO : VERSO),
          image_id:           image,
          image:              image,
          visiblepage:        nil,
          display?:           false,
          toc:                [],
          ill:                []
        }
        @image_data << OpenStruct.new(data)
      end

      @image_data
    end

    def xlsx_path
      File.join @package_directory, @structural_xlsx_base
    end

    def spreadsheet_data
      return @xlsx_data.data unless @xlsx_data.nil?

      @xlsx_data = XlsxData.new xlsx_path: xlsx_path, config: @sheet_config
      @xlsx_data.data
    end

    def xml
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.record {
          xml.ark image_data.first.ark_id
          xml.pages {
            image_data.each do |page|
              data = {
                number:               page.number,
                seq:                  page.number,
                'image.defaultscale': page.image_defaultscale,
                side:                 page.side,
                id:                   page.image,
                'image.id':           page.image_id,
                visiblepage:          page.visiblepage,
                display:              page.display?
              }

              xml.page(data) {
                page.toc.each { |toc| xml.tocentry toc, name: 'toc' }
                page.ill.each { |ill| xml.tocentry ill, name: 'ill' }
              }
            end
          }
        }
      end
      builder.to_xml
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
        next [] if row_hash[:filename].nil?
        next [] if row_hash[:filename].to_s.strip.empty?
        row_hash[:filename].to_s
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
