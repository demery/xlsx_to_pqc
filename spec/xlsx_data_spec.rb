require 'spec_helper'
require 'yaml'
require 'pp'

include XlsxToPqcXml

RSpec.describe XlsxData do
  include_context 'shared context'

  let(:tiff_files) do
    %w{ 0001.tif
        0002.tif
        0002a.tif
        0002b.tif
        0003.tif
        0004.tif
        0005.tif
        0006.tif
        0007.tif
        reference.tif }.freeze
  end


  let(:good_structural_xlsx)    { File.join fixtures_path, 'good_pqc_structural.xlsx' }
  let(:structural_config_yml)   { File.join fixtures_path, 'structural_config.yml' }
  let(:sheet_config)            { YAML.load open(structural_config_yml).read }
  let(:valid_data)              {
    XlsxData.new xlsx_path: good_structural_xlsx, config: sheet_config
  }

  let(:column_headers_xlsx)     { File.join fixtures_path, 'column_headers.xlsx' }
  let(:column_config_yml)       { File.join fixtures_path, 'column_config.yml' }
  let(:column_header_config)    { YAML.load open(column_config_yml).read }
  let(:column_header_data)      {
    XlsxData.new xlsx_path: column_headers_xlsx, config: column_header_config
  }

  let(:missing_required_xlsx)   { File.join fixtures_path, 'missing_required_values.xlsx' }
  let(:missing_required)       {
    XlsxData.new xlsx_path: missing_required_xlsx, config: sheet_config
  }

  let(:fails_uniqueness_xlsx)   { File.join fixtures_path, 'fails_uniqueness.xlsx' }
  let(:fails_uniqueness)       {
    XlsxData.new xlsx_path: fails_uniqueness_xlsx, config: sheet_config
  }

  let(:fails_data_type_integer_xlsx)   { File.join fixtures_path, 'fails_data_type_integer.xlsx' }
  let(:fails_data_type_integer)       {
    XlsxData.new xlsx_path: fails_data_type_integer_xlsx, config: sheet_config
  }

  let(:fails_data_type_ark_xlsx)   { File.join fixtures_path, 'fails_data_type_ark.xlsx' }
  let(:fails_data_type_ark)       {
    XlsxData.new xlsx_path: fails_data_type_ark_xlsx, config: sheet_config
  }

  let(:fails_multiple_xlsx)   { File.join fixtures_path, 'fails_multiple.xlsx' }
  let(:fails_multiple)       {
    XlsxData.new xlsx_path: fails_multiple_xlsx, config: sheet_config
  }

  let(:config_headers) {
    [
      'ARK ID',
      'PAGE SEQUENCE',
      'VISIBLE PAGE',
      'TOC ENTRY',
      'ILL ENTRY',
      'FILENAME',
      'NOTES',
      nil,
      'FILLER'
    ]
  }

  context 'new' do
    it 'should create a XlsxData instance' do
      expect(XlsxData.new xlsx_path: good_structural_xlsx, config: sheet_config).to be_a XlsxData
    end
  end

  context '#headers' do
    it 'should return the expected headers' do
      expect(valid_data.headers).to eq config_headers
    end

    it 'should return the column headers' do
      expect(column_header_data.headers).to eq config_headers
    end
  end

  context '#data' do
    it 'should return an array of hashes' do
      expect(valid_data.data).to be_an Array
      expect(valid_data.data.first).to be_a Hash
    end

    it 'should return an array of hashes when the headers are columns' do
      expect(column_header_data.data).to be_an Array
      expect(column_header_data.data.first).to be_a Hash
    end

    it 'should return the same data whether headers are on columns or rows' do
      expect(column_header_data.data).to eq valid_data.data
    end

    it 'should generate no errors if :data_only is true' do
      fails_multiple.data data_only: true
      expect(fails_multiple.errors).to be_empty
    end
  end

  context '#process' do
    it 'should not extract data if :validation_only is true' do
      fails_multiple.process validation_only: true
      expect(fails_multiple).not_to be_extracted
    end

    it 'should not validate if :data_only is true' do
      fails_multiple.process data_only: true
      expect(fails_multiple.errors).to be_empty
    end
  end

  context '#valid?' do
    it 'should be valid' do
      expect(valid_data).to be_valid
    end

    it 'should be true when required values are missing' do
      expect(missing_required).not_to be_valid
      expect(missing_required.errors).to include :required_value_missing
    end

    it 'should be true when uniqueness fails' do
      expect(fails_uniqueness).not_to be_valid
      expect(fails_uniqueness.errors).to include :non_unique_value
    end

    it 'should be true when not an integer' do
      expect(fails_data_type_integer).not_to be_valid
      expect(fails_data_type_integer.errors).to include :non_valid_integer
    end

    it 'should be true when not an ark' do
      expect(fails_data_type_ark).not_to be_valid
      expect(fails_data_type_ark.errors).to include :non_valid_ark
    end

    it 'should find multiple problems' do
      expect(fails_multiple).not_to be_valid
      expect(fails_multiple.errors).to include :non_valid_ark
      expect(fails_multiple.errors).to include :non_valid_integer
      expect(fails_multiple.errors).to include :required_value_missing
    end
  end

end
