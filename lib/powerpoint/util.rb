require 'htmltoooxml'

module Powerpoint
  module Util

    def pixle_to_pt(px)
      px * 12700
    end

    def render_view(template_name, path)
      view_contents = read_template(template_name)
      renderer = ERB.new(view_contents)
      data = renderer.result(binding)
      File.open(path, 'w') { |f| f << data }
    end

    def render_raw(data, path)
      File.open(path, 'w') { |f| f << data }
    end

    def read_template(filename)
      File.read("#{Powerpoint::VIEW_PATH}/#{filename}")
    end

    def require_arguments(required_argements, argements)
      raise ArgumentError unless required_argements.all? {|required_key| argements.keys.include? required_key}
    end

    def copy_media(extract_path, image_path)
      image_name = File.basename(image_path).gsub('jpg','jpeg')
      dest_path = "#{extract_path}/ppt/media/#{image_name}"
      FileUtils.copy_file(image_path, dest_path) unless File.exist?(dest_path)
    end

    def html_to_ooxml(html)
      source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
      result = Htmltoooxml::Document.new().transform_doc_xml(source, false)
      result.gsub!(/\s*<!--(.*?)-->\s*/m, '')
      result = remove_declaration(result)
      result = remove_whitespace(result)
      result
    end

    def remove_whitespace(ooxml)
      ooxml.gsub(/\s+/, ' ').gsub(/>\s+</, '><').strip
    end

    def remove_declaration(ooxml)
      ooxml.sub(/<\?xml (.*?)>/, '').gsub(/\s*xmlns:(\w+)="(.*?)\s*"/, '')
    end
  end
end
