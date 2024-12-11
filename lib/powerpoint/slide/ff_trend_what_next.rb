require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'sanitize'

module Powerpoint
  module Slide
    class FFTrendWhatNext
      include Powerpoint::Util

      attr_reader :title, :content, :cols

      def initialize(options={})
        require_arguments [:presentation, :title, :content], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}

        format_content = -> (html) {
          case html
          when String
            text = Sanitize.clean(html).strip
            text.empty? ? nil : text
          else
            nil
          end
        }

        @cols = content["rowsManagerInput"]["value"].each_with_index.map do |(row, row_num)|
          columns = row['item']['items']

          (0..2).map do |col_num|
            format_content.call columns["textInput#{col_num + 1}"].dig('value')
          end
        end.transpose

        default_headers = [
          'Previously',
          'Now',
          'In the future',
        ]
        @headers = (1..3).map do |i|
          header = content.dig('headers', 'items', "headerInput#{i}", 'value')&.strip
          header.nil? || header.empty? ? default_headers[i - 1] : header
        end
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        nil
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_what_next_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_what_next_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
