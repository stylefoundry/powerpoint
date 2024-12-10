require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class FFTrendThreeRowText
      include Powerpoint::Util

      attr_reader :title, :content, :links

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :links], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}

        html_guard_blank = -> (html) do
          if ooxml_blank?(html)
            html = <<~HTML
              <p>&nbsp;</p>
            HTML
          end

          html_to_ooxml(html)
        end

        format_text = -> (text) do
          text = text&.strip
          if text.nil? || text.empty?
            "\u00A0"
          else
            text
          end.encode(:xml => :text)
        end

        @main_heading_text = format_text.call(
          content.dig('headingInput', 'value')
        )
        @heading_one_text = format_text.call(
          content.dig('column1', 'items', 'headingInput', 'value')
        )
        @text_one = html_guard_blank.call(
          content.dig('column1', 'items', 'textInput', 'value')
        )
        @heading_two_text = format_text.call(
          content.dig('column2', 'items', 'headingInput', 'value')
        )
        @text_two = html_guard_blank.call(
          content.dig('column2', 'items', 'textInput', 'value')
        )
        @heading_three_text = format_text.call(
          content.dig('column3', 'items', 'headingInput', 'value')
        )
        @text_three = html_guard_blank.call(
          content.dig('column3', 'items', 'textInput', 'value')
        )
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        nil
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_three_row_text_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_three_row_text_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
