# -*- coding: utf-8 -*-
# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
import json
import logging
import textwrap
from dataclasses import dataclass
from email.utils import make_msgid, parseaddr
from typing import Any, Dict, Optional

import bleach
from flask_babel import gettext as __

from superset import app
from superset.models.reports import ReportRecipientType
from superset.reports.notifications.base import BaseNotification
from superset.reports.notifications.exceptions import NotificationError
from superset.utils.core import send_email_smtp
from superset.utils.urls import modify_url_query

logger = logging.getLogger(__name__)

TABLE_TAGS = ["table", "th", "tr", "td", "thead", "tbody", "tfoot"]
TABLE_ATTRIBUTES = ["colspan", "rowspan", "halign", "border", "class"]


@dataclass
class EmailContent:
    body: str
    data: Optional[Dict[str, Any]] = None
    images: Optional[Dict[str, bytes]] = None


class EmailNotification(BaseNotification):  # pylint: disable=too-few-public-methods
    """
    Sends an email notification for a report recipient
    """

    type = ReportRecipientType.EMAIL

    @staticmethod
    def _get_smtp_domain() -> str:
        return parseaddr(app.config["SMTP_MAIL_FROM"])[1].split("@")[1]

    @staticmethod
    def _error_template(text: str) -> str:
        return __(
            """
            Error: %(text)s
            """,
            text=text,
        )

    def _get_content(self) -> EmailContent:
        if self._content.text:
            return EmailContent(body=self._error_template(self._content.text))
        # Get the domain from the 'From' address ..
        # and make a message id without the < > in the end
        csv_data = None
        domain = self._get_smtp_domain()
        images = {}

        if self._content.screenshots:
            images = {
                make_msgid(domain)[1:-1]: screenshot
                for screenshot in self._content.screenshots
            }

        # Strip any malicious HTML from the description
        description = bleach.clean(self._content.description or "")

        # Strip malicious HTML from embedded data, allowing only table elements
        if self._content.embedded_data is not None:
            df = self._content.embedded_data
            html_table = bleach.clean(
                df.to_html(na_rep="", index=True),
                tags=TABLE_TAGS,
                attributes=TABLE_ATTRIBUTES,
            )
        else:
            html_table = ""

        call_to_action = __("Explore in Superset")
        url = (
            modify_url_query(self._content.url, standalone="0")
            if self._content.url is not None
            else ""
        )
        img_tags = []
        for msgid in images.keys():
            #  <img width="1000px" src="cid:{msgid}">
            img_tags.append(
                f"""<div class="image">
                     <img class="adapt-img" src="cid:{msgid}" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="100%">
                </div>
                """
            )
        img_tag = "".join(img_tags)
        #            <html>
        #      <head>
        #        <style type="text/css">
        #          table, th, td {{
        #            border-collapse: collapse;
        #            border-color: rgb(200, 212, 227);
        #            color: rgb(42, 63, 95);
        #            padding: 4px 8px;
        #          }}
        #          .image{{
        #              margin-bottom: 18px;
        #          }}
        #        </style>
        #      </head>
        #      <body>
        #        <p>{description}</p>
        #        <b><a href="{url}">{call_to_action}</a></b><p></p>
        #        {html_table}
        #        {img_tag}
        #      </body>
        #    </html>
        body = textwrap.dedent(
            f"""
                <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
                <html xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" style="font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif"> 
                <head> 
                <meta charset="UTF-8"> 
                <meta content="width=device-width, initial-scale=1" name="viewport"> 
                <meta name="x-apple-disable-message-reformatting"> 
                <meta http-equiv="X-UA-Compatible" content="IE=edge"> 
                <meta content="telephone=no" name="format-detection"> 
                <title>experienz</title>
                <link href="https://fonts.googleapis.com/css?family=Poppins:400,400i,700,700i" rel="stylesheet">
                <style type="text/css">
                        #outlook a {{
                            padding:0;
                        }}
                        .es-button {{
                            mso-style-priority:100!important;
                            text-decoration:none!important;
                        }}
                        a[x-apple-data-detectors] {{
                            color:inherit!important;
                            text-decoration:none!important;
                            font-size:inherit!important;
                            font-family:inherit!important;
                            font-weight:inherit!important;
                            line-height:inherit!important;
                        }}
                        .es-desk-hidden {{
                            display:none;
                            float:left;
                            overflow:hidden;
                            width:0;
                            max-height:0;
                            line-height:0;
                            mso-hide:all;
                        }}
                        [data-ogsb] .es-button {{
                            border-width:0!important;
                            padding:10px 30px 10px 30px!important;
                        }}
                        [data-ogsb] .es-button.es-button-1 {{
                            padding:15px 30px!important;
                        }}
                        @media only screen and (max-width:600px) {{
                            p, ul li, ol li, a {{ line-height:150%!important }} 
                            h1, h2, h3, h1 a, h2 a, h3 a {{ line-height:120% }} 
                            h1 {{ font-size:30px!important; text-align:left }} 
                            h2 {{ font-size:26px!important; text-align:left }} 
                            h3 {{ font-size:20px!important; text-align:left }} 
                            .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a {{ font-size:36px!important; text-align:left }} 
                            .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a {{ font-size:26px!important; text-align:left }} 
                            .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a {{ font-size:20px!important; text-align:left }} 
                            .es-menu td a {{ font-size:12px!important }} 
                            .es-header-body p, .es-header-body ul li, .es-header-body ol li, .es-header-body a {{ font-size:14px!important }} 
                            .es-content-body p, .es-content-body ul li, .es-content-body ol li, .es-content-body a {{ font-size:16px!important }} 
                            .es-footer-body p, .es-footer-body ul li, .es-footer-body ol li, .es-footer-body a {{ font-size:14px!important }} 
                            .es-infoblock p, .es-infoblock ul li, .es-infoblock ol li, .es-infoblock a {{ font-size:12px!important }} 
                            *[class="gmail-fix"] {{ display:none!important }} 
                            .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3 {{ text-align:center!important }} 
                            .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3 {{ text-align:right!important }}
                            .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3 {{ text-align:left!important }} 
                            .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img {{ display:inline!important }} 
                            .es-button-border {{ display:inline-block!important }} 
                            a.es-button, button.es-button {{ font-size:20px!important; display:inline-block!important }} 
                            .es-adaptive table, .es-left, .es-right {{ width:100%!important }} 
                            .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header {{ width:100%!important; max-width:600px!important }} 
                            .es-adapt-td {{ display:block!important; width:100%!important }} 
                            .adapt-img {{ width:100%!important; height:auto!important }} 
                            .es-m-p0 {{ padding:0!important }} 
                            .es-m-p0r {{ padding-right:0!important }} 
                            .es-m-p0l {{ padding-left:0!important }} 
                            .es-m-p0t {{ padding-top:0!important }} 
                            .es-m-p0b {{ padding-bottom:0!important }} 
                            .es-m-p20b {{ padding-bottom:20px!important }} 
                            .es-mobile-hidden, .es-hidden {{ display:none!important }} 
                            tr.es-desk-hidden, td.es-desk-hidden, table.es-desk-hidden {{ width:auto!important; overflow:visible!important; float:none!important; max-height:inherit!important; line-height:inherit!important }} 
                            tr.es-desk-hidden {{ display:table-row!important }} 
                            table.es-desk-hidden {{ display:table!important }} 
                            td.es-desk-menu-hidden {{ display:table-cell!important }} 
                            .es-menu td {{ width:1%!important }} 
                            table.es-table-not-adapt, .esd-block-html table {{ width:auto!important }} 
                            table.es-social {{ display:inline-block!important }} 
                            table.es-social td {{ display:inline-block!important }} 
                            .es-m-p5 {{ padding:5px!important }} 
                            .es-m-p5t {{ padding-top:5px!important }} 
                            .es-m-p5b {{ padding-bottom:5px!important }} 
                            .es-m-p5r {{ padding-right:5px!important }} 
                            .es-m-p5l {{ padding-left:5px!important }} 
                            .es-m-p10 {{ padding:10px!important }} 
                            .es-m-p10t {{ padding-top:10px!important }} 
                            .es-m-p10b {{ padding-bottom:10px!important }} 
                            .es-m-p10r {{ padding-right:10px!important }} 
                            .es-m-p10l {{ padding-left:10px!important }} 
                            .es-m-p15 {{ padding:15px!important }} 
                            .es-m-p15t {{ padding-top:15px!important }} 
                            .es-m-p15b {{ padding-bottom:15px!important }} 
                            .es-m-p15r {{ padding-right:15px!important }} 
                            .es-m-p15l {{ padding-left:15px!important }} 
                            .es-m-p20 {{ padding:20px!important }} 
                            .es-m-p20t {{ padding-top:20px!important }} 
                            .es-m-p20r {{ padding-right:20px!important }} 
                            .es-m-p20l {{ padding-left:20px!important }} 
                            .es-m-p25 {{ padding:25px!important }} 
                            .es-m-p25t {{ padding-top:25px!important }} 
                            .es-m-p25b {{ padding-bottom:25px!important }} 
                            .es-m-p25r {{ padding-right:25px!important }} 
                            .es-m-p25l {{ padding-left:25px!important }} 
                            .es-m-p30 {{ padding:30px!important }} 
                            .es-m-p30t {{ padding-top:30px!important }} 
                            .es-m-p30b {{ padding-bottom:30px!important }} 
                            .es-m-p30r {{ padding-right:30px!important }} 
                            .es-m-p30l {{ padding-left:30px!important }} 
                            .es-m-p35 {{ padding:35px!important }} 
                            .es-m-p35t {{ padding-top:35px!important }} 
                            .es-m-p35b {{ padding-bottom:35px!important }} 
                            .es-m-p35r {{ padding-right:35px!important }} 
                            .es-m-p35l {{ padding-left:35px!important }} 
                            .es-m-p40 {{ padding:40px!important }} 
                            .es-m-p40t {{ padding-top:40px!important }} 
                            .es-m-p40b {{ padding-bottom:40px!important }} 
                            .es-m-p40r {{ padding-right:40px!important }} 
                            .es-m-p40l {{ padding-left:40px!important }} 
                            .es-desk-hidden {{ display:table-row!important; width:auto!important; overflow:visible!important; max-height:inherit!important }} }}
                </style> 
                </head> 
                <body style="width:100%;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0"> 
                <div class="es-wrapper-color" style="background-color:#FAFAFA"><!--[if gte mso 9]>
                            <v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
                                <v:fill type="tile" color="#fafafa"></v:fill>
                            </v:background>
                        <![endif]--> 
                <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FAFAFA"> 
                    <tr> 
                    <td valign="top" style="padding:0;Margin:0"> 
                    <table cellpadding="0" cellspacing="0" class="es-content" align="center" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%"> 
                        <tr> 
                        <td align="center" style="padding:0;Margin:0"> 
                        <table bgcolor="#ffffff" class="es-content-body" align="center" cellpadding="0" cellspacing="0" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px"> 
                            <tr> 
                            <td align="left" style="padding:0;Margin:0;padding-right:40px"> 
                            <table cellspacing="0" cellpadding="0" width="100%" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                <tr> 
                                <td align="left" style="padding:0;Margin:0;width:560px"> 
                                <table width="100%" cellspacing="0" cellpadding="0" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                    <tr> 
                                    <td class="es-m-txt-c" align="left" style="padding:0;Margin:0;padding-top:10px;padding-left:15px;font-size:0px">
                                        <a href="https://experienz.co.uk" target="_blank" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#5C68E2;font-size:12px">
                                            <img src="data:image/png;base64,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" alt="experienz" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" title="experienz" width="300">
                                        </a>
                                    </td> 
                                    </tr> 
                                </table></td> 
                                </tr> 
                            </table></td> 
                            </tr> 
                            <tr> 
                            <td align="left" style="Margin:0;padding-bottom:20px;padding-left:20px;padding-right:20px;padding-top:30px"> 
                            <table cellpadding="0" cellspacing="0" width="100%" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                <tr> 
                                <td align="center" valign="top" style="padding:0;Margin:0;width:560px"> 
                                <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                    <tr> 
                                    <td align="left" class="es-m-txt-l" style="padding:0;Margin:0;padding-bottom:10px"><h1 style="Margin:0;line-height:46px;mso-line-height-rule:exactly;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;font-size:36px;font-style:normal;font-weight:bold;color:#333333"><b>Summary From Last Week</b></h1></td> 
                                    </tr> 
                                    <tr> 
                                    <td align="left" style="padding:0;Margin:0;padding-top:5px;padding-bottom:5px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;line-height:18px;Margin-bottom:15px;color:#333333;font-size:12px"><br></p></td> 
                                    </tr> 
                                    <tr> 
                                    <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px;font-size:0px">
                                        {img_tag}
                                    </td> 
                                    </tr> 
                                    <tr> 
                                    <td align="center" class="es-m-txt-c" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px"><span class="es-button-border" style="border-style:solid;border-color:transparent;background:transparent;border-width:2px;display:inline-block;border-radius:0px;width:auto"><a href="https://portal.experienz.co.uk/insights/organisational-health" class="es-button es-button-1" target="_blank" style="mso-style-priority:100 !important;text-decoration:none;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;color:#FFFFFF;font-size:18px;border-style:solid;border-color:#088A84;border-width:15px 30px;display:inline-block;background:#088A84;border-radius:0px;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;font-weight:normal;font-style:normal;line-height:22px;width:auto;text-align:center">Explore More</a></span></td> 
                                    </tr> 
                                </table></td> 
                                </tr> 
                            </table></td> 
                            </tr> 
                        </table></td> 
                        </tr> 
                    </table> 
                    <table cellpadding="0" cellspacing="0" class="es-footer" align="center" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top"> 
                        <tr> 
                        <td align="center" style="padding:0;Margin:0"> 
                        <table class="es-footer-body" align="center" cellpadding="0" cellspacing="0" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:640px"> 
                            <tr> 
                            <td align="left" style="Margin:0;padding-top:20px;padding-bottom:20px;padding-left:20px;padding-right:20px"> 
                            <table cellpadding="0" cellspacing="0" width="100%" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                <tr> 
                                <td align="left" style="padding:0;Margin:0;width:600px"> 
                                <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                    <tr> 
                                    <td align="center" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px;font-size:0"> 
                                    <table cellpadding="0" cellspacing="0" class="es-table-not-adapt es-social" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                        <tr> 
                                        <td align="center" valign="top" style="padding:0;Margin:0;padding-right:40px"><a target="_blank" href="https://www.facebook.com/experienzcouk" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img title="Facebook" src="https://prmyfz.stripocdn.email/content/assets/img/social-icons/logo-black/facebook-logo-black.png" alt="Fb" width="32" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic"></a></td> 
                                        <td align="center" valign="top" style="padding:0;Margin:0;padding-right:40px"><a target="_blank" href="https://https://www.instagram.com/experienz.co.uk" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img title="Instagram" src="https://prmyfz.stripocdn.email/content/assets/img/social-icons/logo-black/instagram-logo-black.png" alt="Inst" width="32" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic"></a></td> 
                                        <td align="center" valign="top" style="padding:0;Margin:0"><a target="_blank" href="https://www.linkedin.com/company/experienz-co-uk/about/" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img title="Linkedin" src="https://prmyfz.stripocdn.email/content/assets/img/social-icons/logo-black/linkedin-logo-black.png" alt="In" width="32" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic"></a></td> 
                                        </tr> 
                                    </table></td> 
                                    </tr> 
                                    <tr> 
                                    <td align="center" style="padding:0;Margin:0;padding-bottom:35px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;line-height:18px;Margin-bottom:15px;color:#333333;font-size:12px"><a target="_blank" href="https://www.facebook.com/experienzcouk" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"></a>Copyright @ 2022 experienz Limited. All rights reserved.</p><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;line-height:18px;Margin-bottom:15px;color:#333333;font-size:12px">New Broad Street House,<br>35 New Broad Street,<br>London,<br>England,<br>EC2M 1NH<br></p><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;line-height:18px;Margin-bottom:15px;color:#333333;font-size:12px"><a href="mailto:info@experienz.co.uk" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px">info@experienz.co.uk</a><br>44 (0) 203 908 1977</p></td> 
                                    </tr> 
                                    <tr> 
                                    <td style="padding:0;Margin:0"> 
                                    <table cellpadding="0" cellspacing="0" width="100%" class="es-menu" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px"> 
                                        <tr class="links"> 
                                        <td align="center" valign="top" width="33.33%" style="Margin:0;padding-left:5px;padding-right:5px;padding-top:5px;padding-bottom:5px;border:0"><a target="_blank" href="https://experienz.co.uk/" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:none;display:block;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;color:#999999;font-size:12px">Visit Us </a></td> 
                                        <td align="center" valign="top" width="33.33%" style="Margin:0;padding-left:5px;padding-right:5px;padding-top:5px;padding-bottom:5px;border:0;border-left:1px solid #cccccc"><a target="_blank" href="https://experienz.co.uk/privacy" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:none;display:block;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;color:#999999;font-size:12px">Privacy Policy</a></td> 
                                        <td align="center" valign="top" width="33.33%" style="Margin:0;padding-left:5px;padding-right:5px;padding-top:5px;padding-bottom:5px;border:0;border-left:1px solid #cccccc"><a target="_blank" href="https://experienz.co.uk/terms-conditions" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:none;display:block;font-family:'Poppins', 'helvetica neue', helvetica, arial, sans-serif;color:#999999;font-size:12px">Terms of Use</a></td> 
                                        </tr> 
                                    </table></td> 
                                    </tr> 
                                </table></td> 
                                </tr> 
                            </table></td> 
                            </tr> 
                        </table></td> 
                        </tr> 
                    </table></td> 
                    </tr> 
                </table> 
                </div>  
                </body>
                </html>
            """
        )

        if self._content.csv:
            csv_data = {__("%(name)s.csv", name=self._content.name): self._content.csv}
        return EmailContent(body=body, images=images, data=csv_data)

    def _get_subject(self) -> str:
        return __(
            "%(prefix)s %(title)s",
            prefix=app.config["EMAIL_REPORTS_SUBJECT_PREFIX"],
            title=self._content.name,
        )

    def _get_to(self) -> str:
        return json.loads(self._recipient.recipient_config_json)["target"]

    def send(self) -> None:
        subject = self._get_subject()
        content = self._get_content()
        to = self._get_to()
        try:
            send_email_smtp(
                to,
                subject,
                content.body,
                app.config,
                files=[],
                data=content.data,
                images=content.images,
                bcc="",
                mime_subtype="related",
                dryrun=False,
            )
            logger.info("Report sent to email")
        except Exception as ex:
            raise NotificationError(ex) from ex
