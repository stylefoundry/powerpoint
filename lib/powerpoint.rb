require "powerpoint/version"
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/textual'
require 'powerpoint/slide/pictorial'
require 'powerpoint/slide/text_picture_split'
require 'powerpoint/slide/picture_description'
require 'powerpoint/compression'
require 'powerpoint/presentation'

require 'powerpoint/slide/ff_trend_intro'
require 'powerpoint/slide/ff_trend_what_next'
require 'powerpoint/slide/ff_trend_heading_text'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/template/"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/"
end
