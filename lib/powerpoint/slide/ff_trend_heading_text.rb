require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class FFHeadingText
      include Powerpoint::Util

      attr_reader :title, :content

      def initialize(options={})
        require_arguments [:presentation, :title, :content], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        nil
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_heading_text_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_heading_text_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
