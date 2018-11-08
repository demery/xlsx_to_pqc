require 'ostruct'
require 'nokogiri'

module XlsxToPqcXml
  ##
  # Class to generate PQC structural metadata XML for directory containing
  # images and a PQC structural metadata spreadsheet, using a hash spreadsheet
  # configuration.
  #
  # This class validates the spreadsheet based on the configuration and
  # by checking that all files named in the spreadsheet are in the directory.
  # The resulting XML is in PQC structural XML format and also lists files
  # found in the directory but not on the spreadsheet; compare:
  #
  #    <page number="7" seq="7" image.defaultscale="3" side="recto" id="0007"
  #       image.id="0007" visiblepage="4v" display="true"/>
  #    <page number="8" seq="8" image.defaultscale="3" side="verso" id="0002a"
  #       image.id="0002a" visiblepage="" display="false"/>
  #
  # The file +0002a.tif+ is not in spreadsheet (see sample below) and has
  # +@display="false"+.
  #
  # === Usage
  #
  # Given the following YAML configuration, and a directory containing a
  # spreadsheet named +pqc_structural.xlsx+ and files:
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
  #
  # Spreadsheet <tt>./ark+=99999=fk42244n9f/pqc_structural.xlsx</tt>:
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
  # Directory: <tt>./ark+=99999=fk42244n9f/</tt>:
  #
  #        0001.tif
  #        0002.tif
  #        0002a.tif
  #        0002b.tif
  #        0003.tif
  #        0004.tif
  #        0005.tif
  #        0006.tif
  #        0007.tif
  #        reference.tif
  #
  # You can generate the XML for the spreadsheet and directory as follows:
  #
  #     require 'yaml'
  #     require 'xlsx_to_pqc_xml'
  #
  #     sheet_config = YAML.load open('structural_config.yml').read
  #
  #     structural_metadata = XlsxToPqcXml::StructuralMetadata.new package_directory: './ark+=99999=fk42244n9f', sheet_config: sheet_config
  #     puts structural_metadata.xml
  #
  # This will print the following XML:
  #
  #     <?xml version="1.0"?>
  #     <record>
  #       <ark/>
  #       <pages>
  #         <page number="1" seq="1" image.defaultscale="3" side="recto" id="0001" image.id="0001" visiblepage="1r" display="true">
  #           <tocentry name="toc">Pio, Alberto (1512-1518)</tocentry>
  #         </page>
  #         <page number="2" seq="2" image.defaultscale="3" side="verso" id="0002" image.id="0002" visiblepage="1v" display="true"/>
  #         <page number="3" seq="3" image.defaultscale="3" side="recto" id="0003" image.id="0003" visiblepage="2r" display="true"/>
  #         <page number="4" seq="4" image.defaultscale="3" side="verso" id="0004" image.id="0004" visiblepage="2v" display="true">
  #           <tocentry name="toc">Table, f. 2v [=3v]</tocentry>
  #         </page>
  #         <page number="5" seq="5" image.defaultscale="3" side="recto" id="0005" image.id="0005" visiblepage="3r" display="true"/>
  #         <page number="6" seq="6" image.defaultscale="3" side="verso" id="0006" image.id="0006" visiblepage="3v-4r" display="true">
  #           <tocentry name="ill">Decorated initial, Initial P, p. 3</tocentry>
  #           <tocentry name="ill">Foliate design, p. 3</tocentry>
  #         </page>
  #         <page number="7" seq="7" image.defaultscale="3" side="recto" id="0007" image.id="0007" visiblepage="4v" display="true"/>
  #         <page number="8" seq="8" image.defaultscale="3" side="verso" id="0002a" image.id="0002a" visiblepage="" display="false"/>
  #         <page number="9" seq="9" image.defaultscale="3" side="recto" id="0002b" image.id="0002b" visiblepage="" display="false"/>
  #         <page number="10" seq="10" image.defaultscale="3" side="verso" id="reference" image.id="reference" visiblepage="" display="false"/>
  #       </pages>
  #     </record>
  #
  # === Spreadsheet configuration
  #
  # {XlsxToPqcXml::StructuralMetadata} uses a standard configuration as described
  # in the documentation for {XlsxToPqcXml::XlsxData}.
  #
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
    # Create a new {StructuralMetadata} instance. The +package_directory+
    # should contain a spreadsheet named +pqc_structural.xlsx+. +sheet_config+
    # should be a hash as below.
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
    # @param [String] package_directory path to the content package directory
    # @param [Hash] sheet_config sheet configuration as shown above
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
