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


  # Configuration hashes
  let(:structural_config)              { YAML.load open(fixture_path 'structural_config.yml').read }
  let(:lower_case_headings_config)     { YAML.load open(fixture_path 'structural_config_lower_case_headings.yml').read }
  let(:column_header_config)           { YAML.load open(fixture_path 'column_config.yml').read }
  let(:blurg_data_type_config)         { YAML.load open(fixture_path 'blurg_data_type_config.yml').read }
  let(:bad_config)                     { YAML.load open(fixture_path 'bad_config.yml').read }

  # XLSX fixtures
  let(:good_structural_xlsx)           { fixture_path 'good_pqc_structural.xlsx' }
  let(:lower_case_xlsx)                { fixture_path 'pqc_structural_lower_case_headings.xlsx' }
  let(:column_headers_xlsx)            { fixture_path 'column_headers.xlsx' }
  let(:missing_required_xlsx)          { fixture_path 'missing_required_values.xlsx' }
  let(:column_missing_required_xlsx)   { fixture_path 'column_missing_required_values.xlsx' }
  let(:fails_uniqueness_xlsx)          { fixture_path 'fails_uniqueness.xlsx' }
  let(:fails_data_type_integer_xlsx)   { fixture_path 'fails_data_type_integer.xlsx' }
  let(:fails_data_type_ark_xlsx)       { fixture_path 'fails_data_type_ark.xlsx' }
  let(:fails_required_headers_xlsx)    { fixture_path 'fails_required_headers.xlsx' }
  let(:fails_unique_headers_xlsx)      { fixture_path 'fails_unique_headers.xlsx' }
  let(:fails_multiple_xlsx)            { fixture_path 'fails_multiple.xlsx' }

  # XlsxData instances
  let(:valid_data)                     { XlsxData.new xlsx_path: good_structural_xlsx, config: structural_config }
  let(:lower_case_headings_data)       { XlsxData.new xlsx_path: lower_case_xlsx, config: lower_case_headings_config}
  let(:column_header_data)             { XlsxData.new xlsx_path: column_headers_xlsx, config: column_header_config }
  let(:missing_required)               { XlsxData.new xlsx_path: missing_required_xlsx, config: structural_config }
  let(:column_missing_required)        { XlsxData.new xlsx_path: column_missing_required_xlsx, config: column_header_config }
  let(:fails_uniqueness)               { XlsxData.new xlsx_path: fails_uniqueness_xlsx, config: structural_config }
  let(:fails_data_type_integer)        { XlsxData.new xlsx_path: fails_data_type_integer_xlsx, config: structural_config }
  let(:fails_data_type_ark)            { XlsxData.new xlsx_path: fails_data_type_ark_xlsx, config: structural_config }
  let(:fails_multiple)                 { XlsxData.new xlsx_path: fails_multiple_xlsx, config: structural_config }
  let(:fails_required_headers)         { XlsxData.new xlsx_path: fails_required_headers_xlsx, config: structural_config }
  let(:fails_unique_headers)           { XlsxData.new xlsx_path: fails_unique_headers_xlsx, config: structural_config }
  let(:blurg_data_type_data)           { XlsxData.new xlsx_path: good_structural_xlsx, config: blurg_data_type_config }
  let(:bad_config_data)                { XlsxData.new xlsx_path: good_structural_xlsx, config: bad_config       }

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
      expect(XlsxData.new xlsx_path: good_structural_xlsx, config: structural_config).to be_a XlsxData
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

    it 'should return the same data regardless of header case' do
      expect(lower_case_headings_data.data).to eq valid_data.data
    end

    it 'should generate no errors if :data_only is true' do
      fails_multiple.data data_only: true
      expect(fails_multiple.errors).to be_empty
    end
  end

  context 'type validators' do
    it 'should raise an error with an unknown data type' do
      expect {
        blurg_data_type_data.valid?
      }.to raise_error(XlsxDataException) { |error|
        expect(error.errors).to include :unknown_data_type
      }
    end

    it 'should raise an error a bad configuration' do
      expect {
        bad_config_data.valid?
      }.to raise_error(XlsxDataException) { |error|
        expect(error.errors).to include :unknown_data_type
        expect(error.errors).to include :attr_not_defined
        expect(error.errors).to include :no_headings_array
      }
    end

    it 'should not raise an error data_type is added' do
      XlsxData.set_type_validator :blurg, lambda { |v| true }
      expect { blurg_data_type_data.valid? }.not_to raise_error
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

    it 'should skip header validation' do
      fails_required_headers.process data_only: true
      expect(fails_required_headers.errors).to be_empty
    end

    it 'should fail if headers are invalid' do
      suppress_output do # don't print the warning
        fails_required_headers.process
      end
      expect(fails_required_headers).not_to be_extracted
    end

    it 'should not validate headers if :data_only is true' do
      fails_required_headers.process data_only: true
      expect(fails_required_headers.errors).to be_empty
    end
  end

  context '#validate_headers' do
    it 'should say a header is missing' do
      expect(fails_required_headers.validate_headers).not_to be_truthy
      expect(fails_required_headers.errors).to include :required_header_missing
      messages = fails_required_headers.errors[:required_header_missing].map(&:text).to_s
      expect(messages).to match /NOTES/
      expect(messages).to match /PAGE SEQUENCE/
      expect(messages).to match /FILENAME/
    end

    it 'should say a header is not unique' do
      expect(fails_unique_headers.validate_headers).not_to be_truthy
      expect(fails_unique_headers.errors).to include :non_unique_header
    end

    it 'should say the headers are valid' do
      expect(valid_data.validate_headers).to be_truthy
      expect(valid_data.errors).to be_empty
    end

    it 'should accept mixed case headers' do
      expect(lower_case_headings_data.validate_headers).to be_truthy
      expect(lower_case_headings_data.errors).to be_empty
    end
  end

  context '#valid?' do
    it 'should be true when data is valid' do
      expect(valid_data).to be_valid
    end

    it 'should be false when required values are missing' do
      expect(missing_required).not_to be_valid
      expect(missing_required.errors).to include :required_value_missing
    end

    it 'should be false when required values are missing with column configuration' do
      expect(column_missing_required).not_to be_valid
      expect(column_missing_required.errors).to include :required_value_missing
    end

    it 'should be false when uniqueness fails' do
      expect(fails_uniqueness).not_to be_valid
      expect(fails_uniqueness.errors).to include :non_unique_value
    end

    it 'should be false when a value is not a valid integer' do
      expect(fails_data_type_integer).not_to be_valid
      expect(fails_data_type_integer.errors).to include :non_valid_integer
    end

    it 'should be false when not a value is not a valid ark' do
      expect(fails_data_type_ark).not_to be_valid
      expect(fails_data_type_ark.errors).to include :non_valid_ark
    end

    it 'should find multiple problems' do
      expect(fails_multiple).not_to be_valid
      expect(fails_multiple.errors).to include :non_valid_ark
      expect(fails_multiple.errors).to include :non_valid_integer
      expect(fails_multiple.errors).to include :required_value_missing
    end

    it 'should return data after validation' do
      expect(valid_data).to be_valid
      expect(valid_data.data.first).not_to be_empty
    end

    it 'should return data after validation when the columns have headings' do
      expect(column_header_data).to be_valid
      expect(column_header_data.data.first).not_to be_empty
    end
  end

end
