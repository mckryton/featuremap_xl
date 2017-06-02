@d-featuremap @s-backlog
Feature: setup featuremap - add color rule

As an user I want to create an color rule by linking specific tags with border or fill colors to indicate those tags in the featuremap.

  Scenario: add rule
  Given a <tag-name>, <element-type> and <color-target>
  And no rule is using this combination
  When I add those items to the list of color rules
  Then the model visualizer will be executed with this rule

  Scenario: warn for duplicate rule
  Given the list of rules contains a rule that is defined by tag name "@status-done", target "background", color "#000000" (black)
  When I try to add an identical rule
  Then setup form will prevent this
  And I will be warned about a duplicate rule

  @s-backlog
  Scenario: warn for tag names aiming at the same target

  @s-backlog
  Scenario: delete color rule

  @s-backlog
  Scenario: edit rule

  Scenario: choose color
  When I choose the color by form
  Then the resulting color value will be diplayed as hex value

  @s-backlog
  Scenario: enter invalid color value
  When I enter a color value that is not a valid color value
  Then the setup will prevent saving my changes to the rule
