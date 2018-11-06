require 'nokogiri'

module XlsxToPqcXml
  class DescriptiveMetadata

    # --- Default values --- (in case we want to make editable later)
    DEFAULT_PQC_DESCRIPTIVE_XLSX_BASE = 'pqc_descriptive.xlsx'.freeze

    def initialize package_directory: , sheet_config:
      @package_directory     = package_directory
      @sheet_config          = sheet_config
      @xlsx_data             = nil
      @descriptive_xlsx_base = DEFAULT_PQC_DESCRIPTIVE_XLSX_BASE
      @attribute_map         = {}
      @data_for_xml          = {}
    end

    def xlsx_path
      File.join @package_directory, @descriptive_xlsx_base
    end

    def spreadsheet_data
      return @xlsx_data.data unless @xlsx_data.nil?

      @xlsx_data = XlsxData.new xlsx_path: xlsx_path, config: @sheet_config
      @xlsx_data.data
    end

    def attribute_map
      return @attribute_map unless @attribute_map.empty?

      @attribute_map = (@sheet_config[:attributes] || []).inject({}) do |memo,attr|
        memo[attr[:attr].to_sym] = attr
        memo
      end
    end

    def data_for_xml
      return @data_for_xml unless @data_for_xml.empty?

      spreadsheet_data.each do |data_hash|
        data_hash.each do |key,value|
          element = attribute_map.dig key, :xml_element
          next unless element
          next unless value
          ra = value.is_a?(Array) ?  value : [value]
          @data_for_xml[element] ||= []
          @data_for_xml[element] += ra
        end
      end

      @data_for_xml
    end

    def xml
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.record {
          xml.ark data_for_xml['ark'].first
          xml.pqc_elements {
            data_for_xml.each do |key,values|
              next if key == 'ark'
              xml.pqc_element(name: key) {
                values.each do |val|
                  xml.value val
                end
              }
            end
          }
        }
      end
      builder.to_xml
    end
  end
end
