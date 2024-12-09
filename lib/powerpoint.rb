require "powerpoint/version"
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/textual'
require 'powerpoint/slide/pictorial'
require 'powerpoint/slide/text_picture_split'
require 'powerpoint/slide/picture_description'
require 'powerpoint/compression'
require 'powerpoint/presentation'
require 'powerpoint/read_presentation'
require 'powerpoint/read_slide'

require 'powerpoint/slide/ff_trend_intro'
require 'powerpoint/slide/ff_trend_what_next'
require 'powerpoint/slide/ff_trend_heading_text'
require 'powerpoint/slide/ff_trend_three_row_text'
require 'powerpoint/slide/ff_trend_sector_impact'
require 'powerpoint/slide/ff_embeded_slide'
require 'powerpoint/slide/ff_trend_outro'
require 'powerpoint/slide/ff_trend_list'
require 'powerpoint/slide/ff_trend_text_left_chart_right'
require 'powerpoint/slide/ff_trend_text_right_chart_left'
require 'powerpoint/slide/ff_trend_text_right_image_left'
require 'powerpoint/slide/ff_trend_text_left_image_right'
require 'powerpoint/slide/ff_trend_two_column_chart'
require 'powerpoint/slide/ff_trend_two_column_text'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/templates/collision-2025-template"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/collision-2025"
  # TEMPLATE_PATH = "#{ROOT_PATH}/templates/collision-ux2022-template"
  # VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/collision"
  # TEMPLATE_PATH = "#{ROOT_PATH}/templates/nvision-template"
  # VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/new"
end
