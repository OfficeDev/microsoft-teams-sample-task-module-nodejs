// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

// tslint:disable:trailing-comma
export const cardTemplates: any = {
    taskModule: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "{{title}}"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "weight": "Bolder",
                                        "text": "{{subTitle}}",
                                        "wrap": true
                                    }
                                ],
                                "width": "stretch"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "{{instructions}}",
                        "wrap": true
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.ShowCard",
                "title": "Deep Links",
                "card": {
                    "type": "AdaptiveCard",
                    "style": "emphasis",
                    "body": [
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "Button 1 URL",
                                    "value": "{{markdown1}}"
                                },
                                {
                                    "title": "Button 2 URL",
                                    "value": "{{markdown2}}"
                                },
                                {
                                    "title": "Button 3 URL",
                                    "value": "{{markdown3}}"
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Click on the buttons below below to open a task module."
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton1}}",
                            "url": "{{url1}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton2}}",
                            "url": "{{url2}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton3}}",
                            "url": "{{url3}"
                        }
                    ]
                }
            },
            {
                "type": "Action.ShowCard",
                "title": "Invoke",
                "card": {
                    "type": "AdaptiveCard",
                    "style": "emphasis",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Not Yet Implemented"
                        }
                    ]
                }
            }
        ],
        "version": "1.0"
    },
};
