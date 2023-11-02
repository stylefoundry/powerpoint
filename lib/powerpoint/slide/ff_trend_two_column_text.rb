require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

module Powerpoint
    module Slide
    class FFTrendTwoColumnText
      include Powerpoint::Util

      attr_reader :title, :left_col_title, :right_col_title, :left_col_content, :right_col_content, :links

      def initialize(options={})
        require_arguments [:presentation, :left_col_title, :right_col_title, :left_col_content, :right_col_content, :links], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_two_column_text_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_two_column_text_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml

    end
  end
end
