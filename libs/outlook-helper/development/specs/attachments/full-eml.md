---
type: feature
title: Option to return full message .eml body
---
# Return full message .eml body

## Overview

The `get_email` function can, through an optional parameter, return the full email as .eml in a text field. 

## Scope

## User Journeys

## Technical Considerations

The message is returned as a single string, like it comes over the wire to the email server (.eml format).
Headers, bodies, attachments are all included.
