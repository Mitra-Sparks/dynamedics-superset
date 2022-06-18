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
                                    <td class="es-m-txt-c" align="left" style="padding:0;Margin:0;padding-top:10px;padding-left:15px;font-size:0px"><a href="https://experienz.co.uk" target="_blank" style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#5C68E2;font-size:12px"><img src="https://portal.experienz.co.uk/static/brand/logo.svg" alt="experienz" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" title="experienz" width="300"></a></td> 
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
