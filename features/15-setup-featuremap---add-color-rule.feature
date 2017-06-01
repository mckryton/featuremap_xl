@d-featuremap @s-backlog
Feature: setup featuremap - show hide elements

As an user I want to create an color rule by linking specific tags with border or fill colors to indicate those tags in the featuremap.

@s-backlog
Scenario: add color rule
Given a <tag-name>, <element-type> and <color-target>
And no rule is using this combination
When I add those items to the list of color rules
Then the model visualizer will be executed with this rule

@s-backlog
Scenario: warn for duplicate color rule

@s-backlog
Scenario: warn for different rule for the same target

@s-backlog
Scenario: delete color rule 
