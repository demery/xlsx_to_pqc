require 'rspec/expectations'
require 'nokogiri'

RSpec.configure do |rspec|
  # This config option will be enabled by default on RSpec 4,
  # but for reasons of backwards compatibility, you have to
  # set it on RSpec 3.
  #
  # It causes the host group and examples to inherit metadata
  # from the shared context.
  rspec.shared_context_metadata_behavior = :apply_to_host_groups
end


RSpec::Matchers.define :have_xpath do |xpath, opts={}|
  real_count = 0
  match do |actual|
    doc = actual.is_a?(Nokogiri::XML::Document) ? actual : Nokogiri::XML(actual)
    if opts[:count]
      real_count = doc.xpath(xpath).size
      real_count == opts[:count]
    else
      !doc.xpath(xpath).empty?
    end
  end

  failure_message do |actual|
    if opts[:count]
      "expected that #{actual} would match xpath '#{xpath}' #{opts[:count]} times; not #{real_count}"
    else
      "expected that #{actual} would have xpath '#{xpath}'"
    end
  end

  failure_message_when_negated do |actual|
    if opts[:count]
      "expected that #{actual} would not match xpath '#{xpath}' #{opts[:count]} times"
    else
      "expected that #{actual} would not have xpath '#{xpath}'"
    end
  end
end



RSpec.shared_context 'shared context', :shared_context => :metadata do
  FIXTURES_DIR = File.expand_path '../../fixtures', __FILE__ unless defined? FIXTURES_DIR

  let(:fixtures_dir) { File.expand_path '../../fixtures', __FILE__ }
  let(:tmp_dir) { File.expand_path '../../../tmp', __FILE__ }
  let(:staging_dir) { File.join tmp_dir, 'staging'}

  # taken from https://gist.github.com/moertel/11091573
  def suppress_output
    begin
      original_stderr = $stderr.clone
      original_stdout = $stdout.clone
      $stderr.reopen(File.new('/dev/null', 'w'))
      $stdout.reopen(File.new('/dev/null', 'w'))
      retval = yield
    rescue Exception => e
      $stdout.reopen(original_stdout)
      $stderr.reopen(original_stderr)
      raise e
    ensure
      $stdout.reopen(original_stdout)
      $stderr.reopen(original_stderr)
    end
    retval
  end

  def fixture_path file
    File.join FIXTURES_DIR, file
  end
end

RSpec.configure do |rspec|
  rspec.include_context 'shared context', :include_shared => true
end
