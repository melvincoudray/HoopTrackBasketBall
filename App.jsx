import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, Users, LayoutGrid, LogOut, Trash2, ChevronLeft, ChevronRight, ShieldCheck, Plus, X, AlertTriangle, TrendingUp, TrendingDown, Minus, BarChart3, ClipboardList, Download, Camera, Search, Home, Video, Link as LinkIcon, Calendar } from "lucide-react";
import {
  PieChart, Pie, Cell, ComposedChart, Bar as RBar, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
} from "recharts";

// ---------------------------------------------------------------------------
// Config & constantes
// ---------------------------------------------------------------------------

const INK = "#0E1116";
const PANEL = "#161B22";
const PANEL2 = "#1D232C";
const LINE = "#2A323D";
const PAPER = "#F3F1EA";
const AMBER = "#F2A93B";
const AMBER_DIM = "#8A6425";
const TEAL = "#2FB8A6";
const RED = "#E15A4E";

const POSITIONS = { "Point Guard": ["Aiden", "Miro", "Nathan"], "Forward": ["Isayah", "Leo", "Timeo", "Tivon", "Kayode", "Thierry", "Maksim"], "Big": ["Salvator", "Jaspert", "Arthur", "Walter"] };
function positionOf(first) {
  for (const [pos, names] of Object.entries(POSITIONS)) { if (names.includes(first)) return pos; }
  return "";
}
const DEFAULT_ROSTER = [
  "Aiden Bezzola","Arthur Perret","Isayah Juillerat","Jaspert Jansen","Kayode Erogbogbo",
  "Leo Belser","Maksim Jovanovic","Miro Petrusic","Nathan Mettler","Salvator Konate",
  "Thierry Kamtchoua","Timeo Belet","Tivon Jamnik","Walter Njouokou"
].map((n, i) => ({ id: "p" + i, name: n, first: n.split(" ")[0], position: positionOf(n.split(" ")[0]) }));

const FLAG_CH_DATA_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA8AAAAPAAQMAAADAGILYAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAABlBMVEX/AAD///9BHTQRAAAAAWJLR0QB/wIt3gAAAAd0SU1FB+oGBQwMD/aFPC8AAAFlSURBVHja7dxBDQAgDACxJQjAvxsk4WAY4AVL4NEzUAcXIUmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEkf13LbAIPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoMr4Z5lTTAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAw+z34WDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAY/BaWJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSJEmSpIsWqUjAUGphSksAAAAldEVYdGRhdGU6Y3JlYXRlADIwMjYtMDYtMDVUMTI6MTI6MTUrMDA6MDA/Up6OAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDI2LTA2LTA1VDEyOjEyOjE1KzAwOjAwTg8mMgAAAABJRU5ErkJggg==";
const AURORE_LOGO_DATA_URI = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCADwAK8DASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6poorh/GPxR8K+FXkgvtQE96nBtbUebID6HHC/iRSbsJtLc7imuyqpZiAo5JJwBXzF4n/AGhdYu2ePw7ptvYRdBLcfvpD+HCj9ajj8DeO/GXly+LvEL29vMPMEEkpdtvX/VLhR1FQ6iRPPfSOp7rr/wASPCOgkrqGvWYlH/LKFvNf8kzXA6v+0R4dtmKaXp2o379iwWFT+ZJ/Svn3x14VufB+vy6dc/PERvgm24EqHv8AUdCPWvbP2e9Osz4N+2NZ25uzdyKZjGC+BjAz1qHVFFylLl2MHVP2idfmLLp2j6daL2MrPKw/9BH6VLYa98XvFmlpqNjqFtZ6fKrOjxiKPIBIOOC3Y1494n48S6tk/wDL5N/6Ga+lvheptPhPprXIMSrayyMX4wpLEH6Y5qXOQqac203seR+C7vxn8QNUubL/AIS7ULcQxec7PM+CNwGAFI9a1dd8DaXpl4lr4o+Ic8V26Bwkkbn5SSAcljxkGof2befFOq4/58v/AGotdJ8Wdb8L6Z4utk8QeGm1W4+zI3n+eVCpubA29D3/ADqW3ccYx5OaRz2p/BppdGbUvD+vxaknlmVA8eBKAP4XBIzwetcR4H8Haz4vuHTSsRW8RAluZWKome3HJPsK+gPFVnPrfw3B8EXn2O3a28yGG3jVVmjxzGO6nqOO/Brk/wBnjxFp40abQZpEg1BZ2mjVzjzlYDp6kY6elCbsDpx5kuhnXPgHTvDCousfES40+4YZCRsUP12784q5FoHjT7Ab3wV8QJdYgXon2ggk+nJZc+xxXJ/FrwX4gh8XalqS2dzf2V3KZY5oUMhUH+BgORjp6YxW98AvDuv2Gv3Oo3VrcWWmNblHWdSnnNkbcKeeOTn/ABou1rcSV5ctjIg+LnxE0jUDp95eCW7RxG0F7apuDZ6EgA/jmvTYPiR8RNMH/E78Ci6RfvPZyEH8gXrivj61gfF/hxYCh1JSPtG3qE3rs3e/3se1es+PI9Ul8O3KaDqEOn6j5iGOeZwij5hkEkEcjin7RoqMHdq+xkab8fvDrTGDWtP1TSZwcMssO8L+XzfpXfaD478L69gaVrljM56RmQI//fLYP6V5v4f0fWdZ027tPiZY6TewoB5FyhXew7kkY2445GK8Yn8I6TJ8V4vDdrem40qaYKJoXDMgKFtu7kEg8Zq1VfUTco28z7UBBor52Hgzxh4UljXwh40dlOfKsb58B8dgDlT17AVpW/xg8TeGHWHx/wCFpViyF+2WfCn9Sp/BhVqpFlNtfEj3eiuT8H/EHw34twmj6lG1yRk20v7uYevynr9RmusrQadxk8iwwSSyHCIpZj6ADNfn9q96dS1e+v25N1PJOf8AgTE/1r7d+J2o/wBlfD7xBeZwyWUgU/7TDaP1Ir4XXhQPTis5mNbojtvhD4c/4SPxpapKm6ys/wDSZ+MghT8q/i2PwzXsfjKYaf48sNc1PxJYaXp+nxFEtMmSa4DcyAoOgPAHX7oNed/D3xtpHgfwhcPDaXV7rF5IWkdY9sURAIRC5645YgZ61iab4B1XV9U0V9UuorebXXaZc5eby8FmlZew9MnuKwer1Ki+WKS1Z7Z8Q9H0nxt4IS6W8toj5f2ixvJXCKCR90k9AehHY/SvLPhX8S7fwlp02kapZTTwec0kclthmUnAIIJ5GRwQateDdA8O3Pj268L3ltfawtrM6xSTT7YYkRfnOxepLkjsOa1vCV/FrvxOGi6VpemWWh6TcSXINtAFeQx5RS7d/mbP4Uim25KS0Yy+1DwzeXVxrll8OdZvpixleWWFkiLdSxGSD69Kab/xL8Q9LS51GWz0Pwa0hjl8q4WNpAvG0s/UZ47D2NdHd+Krabwv4r8Q+H9Wuft6u8KRX8wKIUxxDHnHIOQeST1rPsPEOg6X4D0PSo/FGkWtxFEslz5lt9sJdvmICjgEMTzzSHbpc8xv7+48BeJ7pfCWqwtHLCq+bFItx8pOdpbaBuyM8DjPWp9J0XxV8UdYN1NL5wiURPezqEjjAyQo2jk8ngevNZ9zpl94s8bypYfaLuO7uhGt59l8tSvA3lQMKMDOK9W8dapb6PZR+APCsz2V6tsGDxkDzG6+UT1DuMnPqQO9UZKN732DSPD/AIe8Ovb6LcePdUN20gRLSzu/KVXY9Nqg45Pc1ir8PvBviKZ/+ET8VuupbiwiuGDEt64IVuvcZrzTwR/yOWidf+P2Lr/vCs6CGe41SOGyDtdPMFhEZw28tgYPY5osHtE1ax6BqviH4g+AL4WF/qU5TGYmlxPHIo7qzDP4dRW5oHiDxt4w0DU9QbxTZ6Ta2TBJnMAQ4IzncoJFb+j3Vj8QvDepeEtQuvtmrabEDHfsB+8kHHmIfQN8pP8AEOe9ch8N/EEHhHQdc0zWpr/SL64mUxTrZNKIyowTgjBpF2s99DmPEHhPXNG8Yafa3RTUNQvnSe3lSQsLglh1Zueo713Hxa8b3914en8P6/4dl0y+uTHMjeesiFVfPYe1aUPi7whqvxGtdde/CLptgyiW4Bj+0THgbVPTALenJHpVPxFqFn4s0Hw5rmghrvUtM1JYTbX7J5jq7ggOAfu7toz6UxWST5WUfg340tdKtrzQfFk/lafKu63N2pKDPDJyPunOfTg1Q0XSrDSfjhpMOi3EFzpctx51s8MgcBGRvlyO4OR9MV3XxJtNW1jw+L6eOSztoruAXWn30MZKYdRmGVeqknn156dKz/iD4I8PeHo9X8RG3uoY0MSWtvpj+UYXIwXY4OAT+n1pFOLSt2Oh+Mkvh2K10oeKbPUJIWkk8q6smw1s2FyTz34/KvOvFdyR4GvV0PxuNZ0aQxrJY3q/6Sg3gjaTzwQM8YroLG58UaP4a0q9HirTr2G/tRcLYa0v3vlBZFkPU88AkVg+Kv7Gu9Bmu9a8E3eh6lLb+da3liCbaRiuV3Y4APuKFoE3e7OR+FmojSfiP4du2O1ReJGx9Ff5D/6FX3KK/POCdreeG4jOGidZFI9Qcj+VfoHpt0l9p9rdRnKTxLKpHcMAf610U2RReh5p+0rffZPhdcwg4a7uYYMeo3bz/wCg18iGvpL9rO926V4esQeZJ5ZyP91QB/6HXgnhbQbvxLrlvpWnmNZ5tx3SEhVAGSTipm9SKl3KyPZvCljYeLPhBosWqTRW9npd5vumchcpGWJH1KsBXN+HfiDaS/ETU/EF9b3U5W2NrplnbRb2VM4Ax0HA592Ndfofwc0vTbLdrt3e6t5f7w2sJMcRb2UHLH8RmsrxF8TrPwGfsmjeAr2wB+UXF5CLSJj/ALwDFvzrJK+x0KErJvQxPh/4f8fWWqX2q6bpUVvc3qsrXGpfLs3NuYhc5zn1FdB4d+DOoWTyPfeJprZrgYlXTwymQZyQWJGRn2rz3Wfiz411cME1K20yFv4bCEbv++3yfyxXJw6rq9tq0GrwatfPq1u3mR3NxO0hz3BBOCp6Ee9Vy+ZmpU46bn0Dqfw78C+DtDuNY1izvr62tFDSM7NKQM4zsXAxzzXJH4s+FdMBHhjwKCR92WdIoM+/Rmr1PwD4osPiF4RaaWFBIym11CzJz5bkYZfdSDkH0PtXy3rujyeHfEOqaJMSWsJ2jRj/ABxnlG/FSKUV3NajVNXgj6e+E/jVvG3h+6uprSKyura5a3lgiYsoGAVIJ65B/SvHfiC2jj4j6nLcXOtJqK3akeTDGVDDbt2ksD6Vo/sz3xh8UeINOJOy5tYrpR/tIxQ/ow/Kun+LvhpdN1OTxrZ28lzPEiqYVXKJMOFnf2UY49QM8ZpNWkKV500zh4IfDkfxSRYJ9UE66rwiwxeUH8zpndnbn2qj4VTw6mvySWtxrJuo4Ll4w8EQwwic5X5vvDkj3ArA8FOX8aaKzMWZr6JixOSSXGTWdbXk+narHeWr7J4JvMQ4zyD3HcdsU7HPzLex6h8DxoyeOY/7Jn1V5jbSBluIY1TbgdSrE9cV6B8XfiG/gmTSLe206DUbi98x3imkKBY0xyDg8knHTtUfwi8IQaWbjxF9mltJdSiXyrSVcG2QnLKPUEgEZ5xjNeTfHq+N78U7qHJMdhZw26jsC2Xb/wBCH5Ukrs6FenSudAfij4H1NCfE3gl7c4y00UMUwHvuXa1dm/wn8Ia7YW9/p8F/pwuI1lj2OVIUjIyj5x9K8T+Gnh0eKfHmmadMm+yhJvLsdjHH0U+zMVH51738YvHo8GaKkGnFH12+BW1jPIiUfelYei9h3P405LZIIWlHmkjlNZ+EWvmFE0rxVLdwQuGjt7x3AVgcjoWXI+grL1BPH+k3Gry+I9FbWbPUrYW10IcMpCjCuNgyGHrivJbLUNVsbx72z1jUre9kYvLPHcsGkYnJLDOD+IrttH+MnjHSAPttxYarbr1+1xeU+P8AfTH6g0cvYzTpyejsdHaeOdNv/CVjoNxfz6Nc2dobaSO7s1uLe4+XHzfxKeODxjNSfD/xDqGlfBzxBeW115l1Z3KCFJz5qqpMYI2n+Hk1r+HfGuk/EdBHq/gDUZc8fa4bYTxD6S/KwqTxF8ELGdHk8N6hNZOw/wBRc5eM+2fvD8c0rWepbhNarU8EnkM08srBQ0jFyFGACTnj2r7Z+Dt9/aPwy8OTk5YWixN9U+T/ANlr4u1fT59K1S70+7Ci4tpWikCnIyDjj2r6p/ZjvvtPw0EBbLWl5LFj0Bw4/wDQq1p7mNLSTR5t+1Xd+b4y0i1DcW9iXI93c/0UVy3wo1KDwxZ+JPFl1EZl061SGGMHBkllcBVB7ZwPwJqz+0XdfafitqCZyLeGGH6fIG/9mqNNIYfs6a3eqvzSajFcE+qRyIv/AMUamWsrFR1qX7GJqfxL8bajdNO2vSWSk5WCyjREQemSCT9Sa2PD/wAYfENkv2bxNBbeIdLfiRXiWObb3xgbX+hA+tcZ4bsI9V8Q6Zp80phiu7mOFpB1UMwGR716T45+Dd5pdvJeeG5pNQt0GXtpFHnKPVccN9MA/Wk2uoRnUa5kO1b4d6B4y0WTxD8LrmOKUEmXTHO1C/dNp5if2+6f1ryIh0mlgnikhuIWMcsMi7XjYdQR61peGdf1Pwnra6pokmy4X5Z7d+I7hB1Rx/I9Qa9c8baFp3xR8KR+MfB6Aa5Am24tuA8wX70Lj/nov8J7/QimU4qqrx3POfhf4obwf41tLyRyum3pW0vl7BSfkk+qk9fQmup/aO0wWfjnTdSQDZqNmY3I6F4m6/8AfLj8q8pOy7tmXB2yAqQeCO2D6EV6h8QNS/4SH4JeCtZuJUa/s7oWs25huYhWiY46nlVJ+tFtRU3zQcWUPgPcG3+K1imeLmzuIj+AVx/6DXffHzxPrXhjXvC82g3ggaSO6E0Uib4plHl/K6nr1PPUZrx3wTrkHhrxzomsXqzNaW0kgmEKb32tGy8DvyRW98XvHFj451fR5NKs76CCxjmVnukVN5fbjABP93vRbW5UJ2pW6nQfDnUtA8X+ObC3vPCyaXqyhrtbmwuyIHaPDcxEcZrN17xBo/g7xNqmneGvC0EupWNw0P8AaGqXBnCuMHcseAO/rUv7PelXl74/OqwRf8S/TreSGeYnA8yRRtQepwMn0FYfxb0m80j4ka0b6IrHqMxvbWQHKyRkAH8QRgj6UrK4O6p81tT2P9n7W9U8QeHdbvdcvXvbz+03TzGGAFEceFUDgAZPArw74iXH2v4leKps8C/aIfRFVf6V1Pwc+I+l+CNN1Ow1mz1Fxc3puUmtolkUKUVcEbgc5X0rzzUL1dR1nVtRRXEd3ez3C7xhtrOSMjtxinbVsKsuamj2P9m+1gstP8U+JL1hHBGVtvMb+GONd7n82H5V5R4l1658V+I77XbzcpuWxBGf+WUA+4v5cn3JruLi+XR/2a9MtLWZPtOvXrLLsYEhWdnYHHT5EUfjXmypLJLDb2kLTXMziKGFB8zuTgKKfUVVtRUETWFpd6nqNvp2l20l3qFwcRQp1PqSeyjuT0r1uPwp4S+GGnW+qeOpE1vxBLzbafGu5A3oiHg47u/HpWnaRab8EvBv2u8SK/8AGOqDaEB++/8AcB6rEnc9z9Rjxa6udS1/XJL3UJJL/V7xwpIXJOeiIo6KOwFK4WjSWurOw1/4ueL9Xfbp9xDoVkOEgtEV3A93YfyAFQaJ8VvGOi3a3FxqbavarzLa3aLll77XUAqfTtXaeFPglLcWP2jxNevaSupK21vgsnHBdjkfgPzrxW6KxRzHcGVA3zDuB3pJpkuVSLTZ3XxfFvc+K4tZsM/Y9Zs4b+Ikf3lwfx+XmvV/2TLvOneI7In7k0UwGf7ylT/6CK87+IujPY/DP4eXEqkSR2At5M9tyK4H/oVdH+yndeX4v1m0z/rrISD6o4H/ALPVQ0YP3apwPxgn+0fE/wASyZz/AKYyf98gL/SvcPCHh9NS+Ctto0gC/b9OcHjo0m4g/mRXzx44nNz4x1+bqXvp2/8AHzX1x4dg+yeH9LtwMeVaxJ9MIKzm9blYdXk2fF2nTXFpLDIf3d5ayDcP7sqNg/qK+0PDmrRa5oVhqdufkuoVkwP4WI+ZfwOR+FfM/wAbPDx8PfEK5mjTbZawpvIiBwJekq/nhv8AgVd5+zr4k3RXfhy5flM3Nrk9QfvqPxw34mnNX1HTfJNxZs/Fb4Yw68kuq6DGkOrgbpIRwtz/AIP79+/rXjPgDxXdeAPFRupklFhKwh1K1IIIAOPMA/vJ+oyK+nPGviW08KaBNqV785HyQwg4Msh6KP5k9hXyV4i1i713WbjUdQHm3l5IAI4kyWY8KiqOvYClHsFW0JJx3O6+OXhWDSNWg8T6PsbRNZIaVo+UjnYZDj/Zcc/UH1rzzQPDt/4gvCPD2jXWpS7iTJEmY0JPd2+Vfzr2/R5IvA3wph0j4jw29+ZnL2ej/wCsdI8hlRyeMK3OegzjnpXnniP4h67rEf2W3mXS9LUbY7KwHlIq+hI5P8varvYVTlTu/uJ1+Gd7aKD4l8S+HtDPeEzG4lH1AwP50Dwr4Fh4u/H+oTMOv2XTto/DINcORltx5Y96DSu2Zc66I9Xv9e8OaP8ADSbwt4I1C/vL27ucu8sDJLKHbLcgAZwAuB2qbWNd8H+M/B/h618VaxqFnq+np801tblm3Y2kElSCCAD9ap+DfD194Eu9M8W+J9PRtLLbPlkDSW5cfJKV6HvxnPNZfi3wlqeoWl54x0/TY7XRLuVp0gEoLxxn+Mj0JycDp6VPU05pWu/uIv8AhE/BVwMWPxCuYXPT7bp3y/mAKST4Y6zOhk8OavoPiFB/BbXPlSn/AIC3H61xPalRjG4eNijjkMpwR+Iq9TPnT3iR6npM+iagItZ0y40u7BJCXEZjDdsg/db8K9b+DOj2Oh6Lf/ETxJ+7tLeN1sA3Xb0aQDuzH5F/+vXP6H8R9Rt7X+zvEcEPiDRX4e2vgHYD/Zc85+ufwrsPidDJ42+H+k3PgXZLoelHddaTGu2ZNq4X5e+wZO3v1GaL30ZrT5b3W55hqV/rHj/xg97JBJNqF43lW1qnIgi/hQemByT65NfRXwy+Hln4RtVubrZc61Ivzz4yIgeqp/U9T9K+d/Aniq48L63BqtgFlQjZLGf+WkZPK57Hjr619baFq1prmkWupadJ5ltcJuU9x6g+hB4NRJsqglKTctzE+J2ujw94J1K8VsTunkQD1d+B+Qyfwr5O03TJNb1bTdHgz5l/cpb/AEUn5j+CgmvUf2gvEg1HxDDots+bbThulx0MzDn8hgfUmm/s5eH/AO0PFF94gnTNvpqG2tyRwZ3Hzkf7q8f8Cpx0VxS/eVLdEegfH2xQ/DkGFMJZ3MJUei8p/UV55+zXceT8U7Zc4E1pPH+gb/2WvYfi1b/afhxrqEZ2weYP+AsD/SvAPg5eGx+I2lT5xxMvX1ieiG4qytUTOWv3NxqtzI3JkuHY/i5/xr7TiXZFGg6KoH6V8TRnddIx7yA/rX24evFKZWG6nC/GTwk3i3wbNHaIDqtiftVme5dRyn/AlyPrivmXwxrU+j6tY6tZ5We2kD7DxnHDIfwyDX194o1/T/DGh3OratN5VrAO3LO3ZFHdieAK+PtZvxrHiHVNVWzjsEvpjMLWPlY89cn+8epxxnNOOqDEJKz6nZfGDxinizX4TYux0y0iURA93YAuT75+X/gNb/wY0Cz0rQ7z4heIE3RW6P8A2dG3ZRwZB/tMflX0H1ryMwS3bxWdv/r7qVLeP2Z2Cj+de/8Ax4eLQvAWi6BYjy7dpFiCjj93EnA/MqaNlYiD3qM8V8Sa3eeItaudT1Fy08zZC54jXsi+wrMrd8J+FtQ8VXc1rpDWpuY03+XNMIyy9yvrjvWrf/DDxlaQmVdG+1xDvaXEcp/LIJ/CmjJRlLWxxtFJLugupLW5jkt7qM4eCZCjr9VPNLTIaa3PTfgzp6a3q9omuairaVZS7rTT55vlmuCOioTyAOTx6etYvxMsv7H1aTTtM1T7X4faV5raCOffHC+fnQqDwyn9PxrkLaeW1uori2kaK4iYOkiHDKR0IqM8sSeSeSaVi+f3bBRRT7CG41G9FnpltcXt2efJtozI31IHQe5pkJN7DK6T4feKrnwj4jgvoWY2rkJdRA8SR9/xHUf/AF607P4VeMrhEaTSorQOcKLm6jUk+mATzXOeJdDuvDuqPp2oPbtdRgGRYZN4QnsT6+1LyL5ZQ947X44+ELfRtTtvEmiqv9jauw81UHyxzsNwceiuOfr9ab8KvH48JabrVpdZkiaI3Foh6efwNv0OQf8AgNdz4atv+Ez+AF1ps3zzwwSwRE9Q8R3Rn8PlFfPVrL51vHJ/eUGj4kazfK1NdSxO95qN+BEr3WpX0+1F6tJK5/xNfXngHw1D4R8J6fo8JDvCm6aUf8tJW5dvxP6Yr5Z8C6/D4T8ZWWuXdkL62hVo3QD54Q3Bkj7FgM8HtnFfXumX9rqmn299p06XFncIJIpUOQynvSlorGuGStfqZnjuPzfBOvoec2M3/oBNfJvhm6az1y0uEOCu7n6ow/rX1v4xx/wiOuZ/58Z//RbV8c2zmORHBwR/hSgRid0KPknGeNr/AMjX23E2+JGH8Sg/mK+LNcgNprGo2+MGG4lTH0YivrGbW1sPh0dczkQaX9qHuRFkfrRIMNpc8A+N3ipvEvjOXT7eQnStGcwoo6SXHSR/fH3R9D61xMGi6hq+j65fWI8qx0m2M11ckcbv4Y1/2jn8B+FZoleHTzPL885UyOT1Z25P5k19I6n4THh39njVtKiQfazpr3N0w6vMQHcn6Yx9BVp2FFe0k5M8Q8Axq/j3wnE3KnUYc574yf6V9C/FnwLf+NZdL+w3ltbJaCTf52453bcYwPY18/8Aw0tb/U/Gfh+50mwuryG1voZZpYoiY41B+Yl+nAPTNfQXxn8ZT+FNCgi01tmo37MkcmM+Ug+8w9+QB9c9qme6Kgl7N8xxuleF/Dfw01m31TxF4lM2pQZaK1to8E5BHKjJI574FM8QaJeeIL9/Enwf8VizvLj95d6WZ/LV37tsPAY9wRgnkGvGZ5ZJ53mnkeWWQlnd2JZie5J60xSVcMpKsOhBwaaujKNZR0S0PYfEcPxE1+/EXiD4d6LqOnqiojXUsayIQoDMJlcEZOTjoKnHwv8ACNl4bm1/xHbXWnxQjc1tY6w1xGxzjaGKg5J4wCfrXjst3cyptlurh19HlYj9TXaeOPESXPgvwloNk4EFvZrcXCqeshJAB+mCf+BUXfQv2yabscXevbS3k0tlZx2VuzfJCjFtq9ssSSx9SasaG+lx6nE2vWJvNPb5JlR2SRV/vIykEEenQ1RopnPzO9z2rWfhR4es9NtdR8OacddS4AaKPUNbaBHyMgLhRuyOxYdKr6KPilp9rf6dongjSdEtZoSkD2TRJ5UmRiQuXO/jI59c1zNr4iS6+EGpaHeOGns7yGS1DddjMSQPoQf++q4oXl0E2LdXITptErAflmkm3udDrJWsj2rwvaW3gO5bXviT4obWPErDbDbrMZ/soPXavqfXAAHA61Tn+HWieNbu71Lwj4oR5JpGmlgukLOjMcnPRgPqDXjXck8k9SasadfXWmXsV5p9xJb3URykkZwR/wDW9qHdkOspaSWh9S/CvwneeEdAuNN1G4t7gy3LSq0OcbSqjByOvFfIl6vl6ffKhK+W8oXHba5x/KvsP4a+KX8V+E4dRlVRexs0NwqDAMigHI9AQQfxr5D8RWV/o8V3aa1YXVhdyGQiO4iK7ixJ+U9G69jRA1qJcseXY1de0PUPDeox2GqgMZoEuLa4UYS4iYAgj0Izgjsa9L/Z18UvYa1N4Vu5M2l2GuLEE/6uUDMiD2YfMB6g+tdh8ZvDAv8A4UwXITOo6HbxXEbY52qiiVfoVyfqor560/Un0jVNM1e3JD2VzFcAjuoYbh+Kk0bkteymmtmfX/jpxF4K15z0FjN/6ARXx/AnmSKgHJ/wr6v+LN0kHw31yVW+WSAIp9d7AD+dfM/grTzqniixswM+Zv8A0Rj/AEpQDEO8ki/8U7T7D8SPEkAGAL13A9m+Yfo1ep318br9mG7kRvnj01oG/wCAybf5Vy/7Sum/YfibJchcJfWsU31IBQ/+gir3w7c6z8FPGeir880MUzxr3w0e4f8Ajyn86clqKlpNo8ciVTc2KuPkNzAGz6eYua+37iGOdJYJkSSGQFGRxlWU8EEdxXwtI7SaaJYfv7BIn1HI/UV9t+HtRj1jQNN1KE5ju7aOcf8AAlB/rSnojTDvRo8a8VfGs6dNc6R4R0JITaSNbma8xHHG6nBCxLycH1Iqh8QbqXx18LPD/i1FX7RZF4b9F6RsSFc47DcoP0YVifHjw8dD8eNqMSYsdaXzgQOFnUAOv4jDfnWl+zzrcVvr2p+GNQCSWWqxm4hjlG5TIoxImDx8yYP/AAGm0rXQSk5SdOR5ZRXuPxC0H4Y6Rqf2TUbO/wBKvXjEqtpyOEYHuBynUHtXlWuWHh+EM+h65d3Q7RXViY2/77ViP/HRQmmc86Tj1MI471b0nT7nVtStbCxTzLq5cRxrnGSfetz4YanpuneLbaTxBbRvp0yPbzCZQ6oHGAxHp/QmvoPRfh9oegzz6r4YgjGoSRn7LLcStLDHuHUAckfjnHek3YdOk5ny9qNldaZdyWuo28ttcxnDRyrtI/z61Y0HRdQ1/UI7LSLWS5nc/wAA+VR6segHua9P8f8AiP4leBJo7nXLfRfEmlTkhJFscLC3904wV9s5z61o/DrU/id4uVdYtzomgaMHxHayWJAuAOpx97HbdkfSqs7XNPq6vueIzRvDNJFKpWSNirA9iDg02vqTW/AHhBdQn8R63bJEV/fXOZituX7sV75Pbv6V83eL9Sg1TxPqV9p9t5NrcXDNEgAUKnQcdumcD1qYu5lUpOG5mUhIAJJAA7mul0XTvCj7X13X9QT1hs7Dn/vtif8A0GvX/hr4e+G+sTTNoel3GoT2gUvNqSO4BOcYDfLnjsO1NtIcKXM7XMGy1a9+GPwetLqBIxres3ge3inXKqpA5Yemxc/8CFa3gj4sW3i/UrTw74l8PqLu6JWNosT27kAsSVblMAZ71578b/EP/CQfEGa3gfdY6Mv2OPHQynmUj6HC/wDAa6T9m3w811rGo+Jrhf3FspsbQn+JzgyMPoML+JosrHRGT5+SOyPb/FKJL4X1hJMeW1lODn08tq+IW50DJ6/Z/wD2WvsH4v6oNI+GniC5BxI9sbeP3eT5B/6F+lfI1xAzWS20Qy77YEHqWIUfzohsTiN0j6N+Md+U+EGlRsf3l59lUj1wm8/qBXC/s8WQvvilYhhlIbeeVv8Avjb/ADYVt/tE3QgPh3RUPFtbmVgPwRf/AEFq0/2UNMMmua7qhXIgt0t1Pu7bj/6AKcFdkVHerY2v2rtH8zStE1lEybeV7WQ/7LjK/qp/OvO/2f8AU0tPGU2nTEeTqVu0RB7svzD9Nwr6U+K2gf8ACS+AdY05F3TmEyw/9dE+ZfzIx+NfFWhalJpGsWOpQEiS1mSUfgeR+IyKqoiZe7NSItZ006BreraTONv9n3MkQ3f8885Q/QqQa+j/AIASX8fwxtBq0T29vDJIbSSU4LW2dysR2Ay2M9gDVPxL8NrHxn490fxN5iNok9qkt5CD/wAfDLgxD6ENhvZMd6oftDeLDp2kW/hbTJBHc6im658vjyrUcbeOm8/L9Aah+9odEI+zbmeY/FLxpJ448Rb7dmGhWLMtinTzD0aY+56D0H1rlLa8udMvrTUdPbbe2Uy3EJ9WU9PoRkfjUJST93DaRGSeVlihiXqzscKo/HFbvjLwhqHgbWIdM1OX7QtxCJ4LkDCucDzE+qtn8CDVehzXlN857/4p0HTvit4P0fV7G9SxbZ58c7R79qkfPGwyOhHrxtrx3UvBWj2I/wCR70SUjqEhkY/+Olq9L/ZptdSi8IX89yf+JTc3RksImHOOkjD/AGWYcD2J71geOfhGtvqd3fWOt6Rp1hPI0kcN85i2E8lQ3QjJP0GKi9nY3nDmXNY8k1CCC2Yra3sV6M43Rxugx6/OBXReE/H3iDwuixadeb7QHP2aceZGPp3X8CKyta0VtLLBdW0W+x2s7l3J/NAP1rNtIZ7uVYoIJZJm+7HGpdj+Aqmjm1i9ND6c+G/xMsfF0n2C6hFlq23d5W7ckoHXYT+eD+tZvxE+LVt4fvJtM0WBL3UIjtlkckRQt/d45Yj04Arnvg98N9Ws9dtNe1yM2Udtl4bdv9Y7EEAsP4Rz0PNYnxL+GWsafrF7qOk28moadPK037obpItxJIK9SAT1FZpK50OdTkOM8VeL9d8TuG1e+eaJTuWBPkiU+yjj8Tk1n6XaWt2yC81OCwz182KRwPxQGqMwkiZkMTb1OCrfKR9c1u6L4cOp4zrvh+zJ7XN06H/0DH61orHMryeup0Wk+AdK1J0SLx3oYkdgqoInDEnoAGK816nepbfB/wCFt40My3OosxWKXZsM9y/CfLk8KOevRTWN8M/hYun65ba1farp2pwQZaFLMl18zsSx44649cVl/tN2up/bdAvZGL6Im+IKv/LO5bozfVRgfQ+tStXY6ox5IuVtTxeFGSP945eQks7nq7E5LH3Jr0v4HeOD4a1tNB1KXGi6hJ+4ZultcMfXsrn8jj3rkvDPgvV/F9hrN9pJ2Q6TEZACM/aZhz5I/wCA5z7lR3rnh5d5aA9UkX8R/wDXFWYRcqbUn1PoH9pttQHh3R1jgc6Qt35l7MvIRgMRBh2XJPPqBXlPwv0k678R9Bs9m+GCb7bN7JEMjP1baK9w+DviSLxx4Gm0vXQl1fWa/Y72OTnz4yPkkI/2l6n1BrK8L+D4vhRYeMNeuLgTgL5Wnsxywh6op/2i7AH/AHRUp2Vux0yipSU+h5j8XtW/tj4g6rMjboYHFrGfZOD/AOPbq+g/2aNG/s34cLeSLiXUbh7jPfYPkX/0En8a+UrO3uNS1CC2i3SXd3Msa9yzu2P5mvvbw9pcOi6FYaZbDENpAkK++0AZ/GrpI56fvScjQbpXxR8Z/DR8L/EHUraNNtpdN9rtsDA2OScfg24fhX2xXkX7SHhI674PXVbSPdfaSTKcDloT98fhgN+BrSaui6kbo534BeIv7S8JSaZOS91pjbVUdWiblcfQ5H5V8+a/rFz4i8R6nrN+rLcXU7Dy26wop2rH/wABAwffNbfw38St4W8W2d+zH7I58m5A7xN1P4cH8K2/jn4YXQvFqaxZKv8AZOt/vQyfdS4xlvwcfMPfNYRWrQm3OnZdA+AWgDWvHzahOu610WITAHoZ3yE/IBj+Ve4/E3wXa+OfDTabPJ9nuEcS21yBkxP0P4EZB+vtXM/s46YLT4enUCuJdTupbgn1RTsT9FJ/GvU6mTs9Oh1UoWhY8w+Muqf8IV8MYNP0BzaTytDp9mU4MSjlmH0VT+dVNOu7P4x/DeSCfyoNfsyBIv8AzxnA4Yf7DjP5+1cj+0tqRuPFOh6Up/d2lq924/25G2r+in86828OeI77whrcWt6WcyRDZPAT8txFnlG/oexqkroxnUSnyvYrX1pPp95PaXsTQ3MDmOSNhgqwr2/4XQT+H/g/rniPS7eOXV5Enmh3ruyIgQq49Mhjj3re+Kfw8TxdZx6tpaLb62sQJRsATrjOxj2YdAfwPtpfBeCe2+HtpaX9u8E8M08UkUq4IPmNwQfrUuRNOk4VLM86vvFvxutbl4V8L2VwFxiWC1LI2QDwd/vUa+NfjSltd3Fz4bsbWG1gad3mtWUEKOQPn5b0FfQlIVDAgjIPBB70+ddjp5fM8N/aK0pDYaHrXlKl2/8Ao05Ufeyu8Z+hDfnXkfh3RbzxBrFtpmnR7ridsA9kXuzegAr6B+O+m32taLomnaVbPc3M99kKg6ARtyT0A56mpPD/AIat/hf4H1fVxAuoaxDaPPOy8Btoz5ansoPfqfyqU9LHNOlzVPIxviVr1t8N/A1p4V8NOF1i6hKI4+9Eh+/O3+0STj3+lbuiJF8TvgvbQXzAT3lp5LyHkx3MZ27/AKh1B/GvmrUNQvNa1K51bVZvtF/dtvkk7AdlUdlA4Ar2/wDZi1IvpfiDSGP/AB63S3MY/wBmVef/AB5D+dW1ZF06ilNpbHqPgvw7aeE/DVjo9gMx26/PIRzLIeWc+5OTXzB8UPD48M/ELVbGFdlnckX1sAOAkhO5R9HDD8q+uq8Q/ac0wGy8P6yi/PDcNZSN/syLuXP/AAJP1qYPX1LrQvDQ82+FOtXOg/EXSJLNHlS/kFhcQp1kRuh+qkBvpmvSv2jfEGI7Dw9bvyx+1XIHp0QfzP5Vh/AbRLe3Gp+OtZwmn2Ebw2bMOpA/eyD/ANAH1NeeeJtZuPEOv3uqXQPm3Um4J12r0VR9BgVT1ZzuThT5e56P+zX4aOs+ODqkyZtNJTzORkGVgQg/Abj+Ar60HSuE+DHhL/hEPA1nbToF1C5/0m69Q7Dhf+AjA/A13dbxVkVTjyxCmyoskbJIqsjAqysMgg9QadRVFnxR8YPBj+C/F89rGh/sy5zPZt22E8p9VPH0x612fgC5tviH8Pb/AMFazKFvraMNZzH7yqPuMPdDgH/ZNe4/FfwXB438KzWJ2JfRfvbOY/wSAdD/ALJ6H8+1fHNhd6n4T8SJOiPbalYTFXjcYwQcMrD0PI+hrCcbO6Mf4cr9D6r8AaTJoPgnRNKuEVJ7S1SKUKcjeB82PxzXQVjeEfEFn4n0G31SwOEkGJIyeYnHVT9P1GK2RyRWL8zujZrQ+TfjFem++K2vtnK23k2q+22ME/qxrlbG2+26tpdnjP2m9ghI9Q0i5/Srviq4+1+M/ElznPm6ncEfQOV/pVrwBCLj4jeFImzj+0Y3/wC+QW/pWx571qH0d8WvEGpeFtHsNV0tVeOK8VLiJx8skbKwwfTnHPrilsde0z4j+E7210i/eyv5YuVz++tZBgq+O4DAHI4OK86+LPivU9A8V6toGs24vvDurwLNbI52shwA3lvjgq4zg5HTpmvGrO91C3muLvRor1msV8y4urbKi3QnGWYdKzUGbyqSU7I9q1bTfi9BdOZ/HPh20yeELpGPTIDR59/xqTSNE+LOoO8MvjnRJ7SZTFK9uyyPGp6sm1B82M45rxGRmmdpJWaSRjks5ySfcmnQ3c2nMbu1mkt5YgWEkTFGGPcVepH1jXY+pfGnjnR/AmmRWSt9q1COJY4bRWywAAALn+EcfU9qd4Tmv/E/wukm1oq9zqVvc5AXaArbwoA9MYxXyy73K6mieILS9t7mXbO8c+VkmRudyseuR3r3v4NeKtX8Y+J7t1iSz8N6VbC1itoR+78wkADcfvFVU8+445qOVo0hOUpWZ88aaS1hAW67AD+HFep/s6Xv2b4jXloThb3TmOPVo3B/kxrzNY/IluoB/wAsbiaP8pGFdh8HJ/s/xY8PtnHmieA/8CiJ/pWjW5hT0qH1lXHfFjwvP4x8Hvo9q4imkuYHEp/5ZhXG5vqF3V2Ncp8SPF8Hg/w+90dr382Y7WI/xPj7x/2R1P4DvWK30O6TSWp5T8Z9btNI0qw8DeH8R2VlGn2kL7fdQ+p/iPuRVf8AZ68EnxJ4pGq3sW7S9LYP8w4lm6qvvj7x/D1rz/Q9L1Pxd4lhsbUtcajfSlmkfnBJyzsfQck19teCvDdl4U8OWek6cv7qBfmcj5pHP3nPuTW0IHEv3kuZm6BxRRRW5sFFFFABXiX7QHwzbXbd/EehQ51W3T/SYUHNzGO4/wBtR+Y47CvbaRhkUmriaTVmfEXwx8az+Dta8xg8mmXBC3UI647Oo/vD9RxX1Vp93b39rb3dlMk9tMoeORDkMD3ryn47fCZt9x4k8L2+c5kvbOMc+8iD+aj6jvXnvwp+IkvhO6Flfl5dEmbLAcm3Y/xqPT1H49evPOFiKVR03yy2PPppDNfX8rfeku52P4yNXTfClQ3xU8LA/wDPxI35QvXUfFL4cGzWTxR4MH23Q7kGee2gO8w55MkePvJ3K9R/LlPhLMj/ABR8KyRsGRriQAjpzC9PfVE8jjUTPY/Hlz4a8b+LL34e+IP9E1CKOO4029U/N5jJkhffp8v8Q9wK2PCHw5h0T4XXnhO4kilnvY51uLhFx5jvkKxzzwNv5V4D8fbZpPjbfFZZIXW3glSRDhlIjXBB+or0j4f/ABmeC1isfGKSTyJhEv4Fyz/9dE9fcdfSk1ZKx0+0ipWkeH2yz28k1jeoY76zcwTxt1VlOP6Vc0jS5fEfiTTNAtQWlvZ1WTH/ACziBy7H8Aa+i/id8HdK8a3/APaltdy6TrBAElxEm5ZgOBvXI5x3B+uateAPh3ofwx0y+1Iyy3t8IWe5vpV+by1G4qi9hx9T60cy3M1QtK/Ql+LPw5i8d6bpFpBOllJZXAIuAmXWDaQyL7n5cduKg+GWueHLXxDd+B/CUKmy0e2EktyG3ebMXCvz/EfVvXgcCvO/iV8XLzXbKXS/DAl06ymXbLdt/r5F7qoHCA+vJ+lZf7KNt5Hj/wAQIhJSOxC5PvItCV4u7NFOMpWicRrC7PEOuIB93Urof+RWrW+HUhh+JXhRx1N+qf8AfSsP61j69Mo8Q645yd+p3IVVGSxMrAADuTXrnw38CQeFrVfGnj5haNbDzLSyY8wk9GYfxSHsvb69KbtqzmjFuo30R7X4i1qy8P6Rcalqcojt4R+LnsqjuTXyl4t8Q6j418SG7ljd5JWENrax/NsUn5UUdyf1NXviB4zv/G+sphJEskbZaWi/McngEgdXP/1hXu/wO+FA8Nxxa74hiVtakXMMJ5Fop/8AZz3Pbp61MINlVJuq7LY2vgj8OI/BWjm71BFbXbxQZ26+SvURA+3c9z9BXp1FFdCVi0klZBRRRTGFFFFABRRRQAhGa8H+MXwXXUnm1vwhEkd6cvPYrhVmPdk7BvbofY9feaCKTVyZRUtz4q+H/j3U/BF+1ndRyy6d5mJ7KTKvE3cpn7reo6H9a9NsfA/hzxB4m0bxt4Lu4rdoLpZ7q1QYjkJBDZT/AJZyYJ9j+teifEz4V6P42RrjaLHVwuFvIl+96CRf4h79R618uWt3q/w98Z3FvHcvDcWVx5NysZOyZVPIIPUEcjI71g48uxKk6bs9Ubv7SenG1+JekaiB8l9YGLPq8ZP9CtedHOOOD2r6H/aG0Qa/8PINb04ebNpTrfRled0LAb/0w3/Aa+d43EkauhyrAEfjSWqJxCtK59m+FdSj1jw1peoRNuW4t0Y+zYww/Ag1hfF7VE0v4eau7Nh7iP7LGO5Z+P5ZP4V5d8EPH1vowfQtcnWGxkYvbXDnCxMeqsewPXPY/Wsb4x+OE8V6tHaaaxOk2ROxunnSdC/07D8fWoSdzaVZez8zzqvY/wBlTTyJPFWsEfLJPHaofXaCzfzWvFrycW9tJKeqjgep7Cvp7wJp3/Cuvg0sl4Al3FbPezg/89nGQp9wSq/hVvRGWHVm5GDpvhfwz8LBN4h8TXC6j4guJZJoEAz5ZZiSIlP15c/pXmvifxFr3xG8QQQxwyysz7bWwtwWCe/ufVj+gqLwxpOrfEfxrBZT3zPe3W55bmfLbEUZJwO3YDgZIr60+Hvw/wBF8EWJi0yHzLuQfvryXBlk9s9h7CnGLk7sG3U0WiOU+D/wjtvCSR6prQjuteYZXHKWuey+rerflXrYGKKK3SsaJJKyCiiimMKKKKACiiigAooooAKKKKAAjNfLP7UHhl9P8VW2uwp/ouoxiOQgdJkGP1XH/fJr6mrm/iD4XtvGHha90i6IRpV3Qy4yYpB91vz6+xNTJXRM43R5D8C/EkOu+Fn0G/KyXNihj2Pz5tueB9cZKn2xXiHxA8JT+BfFUumMrHS7gtNp0x6MmeYyf7y5x9MHvS202seA/GBJRrbVNPlKSRt91x3B9VYd/Qg19DTxeH/i74FKSZUHByuPOspwOCP84IrnT5XqKDVWPI90fLdFa/irwvq/hHVDp+tRdSfIu4wfKuV9VPY+qnkVo+AvAmp+N74Jb77TSI2xc35HAHdI/wC8/wCg7+lUYKnLm5TQ+C/gxvF3ipNRu4idC0mQOxPS4uByqD1A6n8B3ruv2hvFK+XB4btJAXJE95jsP4EP/oX5V1/irXdF+F3g+30/SYIklWMx2VpnJJ7yP3xnknua8C8L6Hqvj7xglpFI8l3dyGW5uWGRGufmdvp2HrgVPxPQ3m1Bckdz2f8AZW8NNHb6l4kuEwJv9EtsjqoOXYfjtH/ATX0HWf4f0m00LRrPTNPj8u1tYxHGvsO59zyT7mtCumKshxjyqwUUUUygooooAKKKKACiiigAooooAKKKKACiiigDzH4zfDGDxtYi80/ZBr1uuIpDwsy/883P8j2+lfMWh6xrngLxNIUjktL2E+Xc2k4wHH91h/I/iK+68VxnxE+HWi+OLMLqEZhvkGIb2EASJ7H+8vsfwxWc4XM5Q15o7nIeGfE3h34jaM9tNBDK+AZ9PugGZD/eHqPRh+lVPHnj7SPA9gun6fFDLqKpthsoQFSEdi+Puj26n9a8d8Y/DfxX4DvftqJLNaxHdHqNln5P94DlD+nvV7wH8IPEfi6ZbzUFk0zTpDva5uQTLLnqVQ8kn1OB9ax5HsHtpvS2pydtb6/8QfFeyNZL/Vbpsk9FjX1PZUH+ea+t/hb4AsfAuiC3hKz6hPhru624Mjf3R6KOw/GtPwT4M0bwbpn2TRbbYWwZZ35kmPqzf06CujreMOUIwtqwoooqzQKKKKACiiigAooooA//2Q==";

const DEFAULT_TEAMS = [
  { id: "u16-sui", name: "U16 Swiss National Team", tagline: "Swiss U16 National Team", logo: FLAG_CH_DATA_URI, logoBg: "#FF0000" },
  { id: "aurore-vitre", name: "Aurore Vitré Basket Ball NM1", tagline: "Aurore Vitré Basket Ball · NM1", logo: AURORE_LOGO_DATA_URI, logoBg: "#F3F1EA" },
];

// Liste des équipes — désormais modifiable depuis le panneau Admin (ajout/suppression),
// stockée globalement (clé non préfixée, lue avant toute sélection d'équipe).
async function loadTeams() {
  const stored = await storeGet("app_teams");
  return stored && stored.length ? stored : DEFAULT_TEAMS;
}
async function saveTeams(teams) { await storeSet("app_teams", teams); }

function simpleHash(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) { h = (h << 5) - h + s.charCodeAt(i); h |= 0; }
  return String(h);
}

function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
// Volontairement PAS new Date().toISOString().slice(0,10) : ça convertit en UTC et peut
// décaler la date d'un jour pour un fuseau horaire en avance sur UTC (ex. la Suisse),
// surtout en fin de journée locale. On utilise les composants de date LOCALE à la place.
function todayLocal(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function matchKey(date, opponent) { return `${date}||${opponent}`; }

// Saison basket = juillet à juin. Sert de valeur par défaut, modifiable dans Settings.
function defaultSeasonLabel(d = new Date()) {
  const y = d.getFullYear();
  return d.getMonth() >= 6 ? `${y}-${y + 1}` : `${y - 1}-${y}`;
}

// Télécharge le contenu déjà rendu d'un élément (id donné) sous forme de fichier HTML
// autonome. Contrairement à window.print(), le téléchargement de fichier n'est pas
// bloqué par le bac à sable de l'iframe — c'est le même mécanisme que la Sauvegarde.
// Le fichier obtenu s'ouvre dans n'importe quel navigateur normal, où l'impression en
// PDF (Ctrl+P) fonctionne sans restriction.
function buildReportHtml(elementId, filename, theme = "light", orientation = "portrait") {
  const el = document.getElementById(elementId);
  if (!el) return null;
  const pageRule = `@page { size: A4 ${orientation}; margin: 16mm; }`;
  const lightStyle = `
  ${pageRule}
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; background: #fff; padding: 24px; max-width: 900px; margin: 0 auto; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 12px; }
  td, th { border: 1px solid #ccc; padding: 5px 8px; text-align: left; font-size: 13px; }
  h1, h2 { color: #111; }
  /* Certains blocs réutilisent les composants sombres de l'appli (ex. comparaison par poste) —
     on force le même passage en clair que dans l'aperçu d'impression de l'appli. On exclut
     les petits éléments dont la couleur EST l'information (pastilles de légende, mini-barres
     de fréquence) pour ne pas les rendre invisibles (blanc sur blanc). */
  .printable-root *:not(.keep-color) { background: #fff !important; color: #111 !important; border-color: #ddd !important; }
  svg { overflow: visible; }
  img { max-width: 100%; }
  [data-new-page="true"] { page-break-before: always; }`;
  const darkStyle = `
  ${pageRule}
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; background: ${INK}; color: ${PAPER}; padding: 24px; max-width: 900px; margin: 0 auto; }
  h1 { color: ${PAPER}; }
  svg { overflow: visible; }
  img { max-width: 100%; }
  [data-new-page="true"] { page-break-before: always; }`;
  const styleTag = theme === "dark" ? darkStyle : lightStyle;
  const full = `<!doctype html><html><head><meta charset="utf-8"><title>${filename}</title>
<style>${styleTag}</style>
</head><body>${el.outerHTML}</body></html>`;
  return { full, bodyHtml: el.outerHTML, styleTag, filename };
}

// Tente un téléchargement de fichier — peut échouer silencieusement selon l'environnement
// (iframe imbriquée). Ne pas se fier uniquement à ce mécanisme : toujours proposer aussi
// l'aperçu à l'écran (ExportModal) qui, lui, ne dépend d'aucune API de fichier.
function tryDownload(full, filename) {
  try {
    const blob = new Blob([full], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    return true;
  } catch { return false; }
}

// Génère un vrai fichier .pdf (pas un .html) à partir du contenu d'un élément, via jsPDF +
// html2canvas — bibliothèques disponibles uniquement sur le site déployé (Netlify), pas dans
// l'artifact Claude. Si l'import échoue (contexte Claude), renvoie null pour que l'appelant
// retombe sur l'export HTML existant, qui lui fonctionne partout.
// Réglages volontairement choisis pour un fichier LÉGER, encore réduits sur retour de
// l'utilisateur (poids toujours trop lourd malgré un premier passage) : scale 1.15 (au lieu
// de 1.5, lui-même déjà réduit depuis 2) et JPEG à 72% de qualité (au lieu de 88%) — le texte
// reste lisible pour un rapport à consulter à l'écran ou imprimer, sans viser une netteté
// "impression photo" qui n'a pas lieu d'être ici.
// Pagination "intelligente" : les éléments marqués data-no-split (cartes de graphiques) ne
// sont jamais coupés par un saut de page — si une coupure "naturelle" tomberait au milieu
// d'un graphique, on la déplace juste avant lui, quitte à créer une page plus courte (et donc
// plus de pages au total, ce qui est très bien) plutôt que de couper un graphique en deux.
// Les éléments marqués data-new-page="true" démarrent TOUJOURS une nouvelle page (même s'il
// restait de la place sur la précédente) — utilisé pour garantir une page par joueur.
async function tryExportPdf(elementId, filename, theme = "light", orientation = "portrait") {
  const el = document.getElementById(elementId);
  if (!el) return false;
  try {
    // BUG RÉEL CORRIGÉ : la bibliothèque jsPDF (v2.x) exporte "jsPDF" en export NOMMÉ, pas en
    // export par défaut — importer "{ default: jsPDF }" donnait silencieusement "undefined",
    // ce qui faisait échouer "new jsPDF(...)" et basculait systématiquement vers le repli
    // HTML, même quand la bibliothèque était correctement installée. C'est pour ça que l'export
    // donnait toujours du HTML au lieu d'un vrai PDF.
    const [{ jsPDF }, { default: html2canvas }] = await Promise.all([
      import("jspdf"),
      import("html2canvas"),
    ]);
    const bg = theme === "dark" ? INK : "#ffffff";

    // Repère chaque élément "à ne pas couper" et chaque élément "nouvelle page forcée" AVANT
    // la capture, en pixels DOM relatifs au haut de l'élément exporté.
    const elTop = el.getBoundingClientRect().top;
    const noSplitEls = Array.from(el.querySelectorAll('[data-no-split="true"]')).map(node => {
      const r = node.getBoundingClientRect();
      return { top: r.top - elTop, bottom: r.bottom - elTop };
    });
    const newPageEls = Array.from(el.querySelectorAll('[data-new-page="true"]')).map(node => {
      const r = node.getBoundingClientRect();
      return { top: r.top - elTop, bottom: r.bottom - elTop };
    });

    const canvas = await html2canvas(el, { backgroundColor: bg, scale: 1.15, useCORS: true });
    const imgData = canvas.toDataURL("image/jpeg", 0.72);
    // A4 en points (72dpi) — inversé en paysage (largeur/hauteur permutées).
    const pageWidth = orientation === "landscape" ? 841.89 : 595.28;
    const pageHeight = orientation === "landscape" ? 595.28 : 841.89;
    const imgWidth = pageWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    // Conversion des rectangles DOM (px) vers l'échelle de l'image finale dans le PDF (pt) :
    // même ratio que celui utilisé pour convertir la hauteur totale de l'image.
    const domToImgRatio = imgWidth / el.getBoundingClientRect().width;
    const noSplitRanges = noSplitEls.map(r => ({ top: r.top * domToImgRatio, bottom: r.bottom * domToImgRatio }));
    const newPageRanges = newPageEls.map(r => ({ top: r.top * domToImgRatio, bottom: r.bottom * domToImgRatio }));

    const pdf = new jsPDF({ unit: "pt", format: "a4", orientation, compress: true });
    let pageTop = 0; // position (en pt, dans l'image totale) du haut de la page courante
    let firstPage = true;
    while (pageTop < imgHeight - 0.5) {
      let pageBottom = Math.min(pageTop + pageHeight, imgHeight);
      // Si un élément protégé chevauche cette limite de page (commence avant, finit après),
      // on raccourcit la page pour s'arrêter juste avant lui — il passera entièrement sur la
      // page suivante. On ignore les éléments plus grands qu'une page entière : rien à faire
      // dans ce cas, ils seront coupés de toute façon faute de place, mais c'est un cas limite.
      for (const r of noSplitRanges) {
        const spansBreak = r.top < pageBottom - 0.5 && r.bottom > pageBottom + 0.5;
        const fitsOnOnePage = (r.bottom - r.top) <= pageHeight;
        if (spansBreak && fitsOnOnePage && r.top > pageTop + 0.5) {
          pageBottom = r.top;
        }
      }
      // Une "nouvelle page forcée" qui commence APRÈS le haut de la page courante coupe la
      // page ici, même s'il restait de la place — chaque joueur démarre ainsi toujours une
      // page fraîche, plutôt que de s'enchaîner à la suite du précédent.
      for (const r of newPageRanges) {
        if (r.top > pageTop + 0.5 && r.top < pageBottom - 0.5) {
          pageBottom = r.top;
        }
      }
      if (!firstPage) pdf.addPage();
      pdf.addImage(imgData, "JPEG", 0, -pageTop, imgWidth, imgHeight);
      // Masque tout ce qui dépasse de la page courante avec un rectangle de la couleur de
      // fond, pour que le contenu des pages suivantes n'apparaisse pas en dessous.
      if (pageBottom < imgHeight - 0.5) {
        pdf.setFillColor(bg);
        pdf.rect(0, pageBottom - pageTop, pageWidth, pageHeight - (pageBottom - pageTop), "F");
      }
      pageTop = pageBottom;
      firstPage = false;
    }
    pdf.save(filename.replace(/\.html$/, ".pdf"));
    return true;
  } catch (e) {
    console.warn("[pdf] jsPDF/html2canvas unavailable, falling back to HTML export:", e);
    return false;
  }
}

function ExportModal({ report, onClose }) {
  if (!report) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 1000, padding: "24px 12px", display: "flex", flexDirection: "column" }}>
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: "12px 12px 0 0", padding: "12px 20px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", maxWidth: 900, width: "100%", margin: "0 auto" }}>
        <div style={{ fontSize: 12.5, color: "#D8DCE2" }}>
          Exact preview of the file — same content as the downloaded one. From this window: print (Ctrl/Cmd+P) or
          select all (click inside then Ctrl/Cmd+A) and copy.
        </div>
        <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
          <button onClick={() => tryDownload(report.full, report.filename)} style={{ ...btnSecondary, fontSize: 12 }}>Retry download</button>
          <button onClick={onClose} style={{ ...btnPrimary, width: "auto", padding: "8px 14px", fontSize: 12 }}>Close</button>
        </div>
      </div>
      <iframe
        title="Report preview"
        srcDoc={report.full}
        style={{ flex: 1, width: "100%", maxWidth: 900, margin: "0 auto", border: "none", borderRadius: "0 0 12px 12px", background: "#fff" }}
      />
    </div>
  );
}

// Export/import complet des données d'une équipe — contourne le préfixage automatique de
// storeGet/storeSet pour parcourir TOUTES les clés (roster, matchs, box scores, scouting,
// objectifs, entraînement, évaluations, photos...), quelle que soit leur nature.
async function exportAllData(teamId, teamName) {
  const prefix = "team_" + teamId + ":";
  const data = {};
  try {
    if (window.storage) {
      const r = await window.storage.list(prefix, true);
      const keys = r ? r.keys : [];
      for (const fullKey of keys) {
        try {
          const got = await window.storage.get(fullKey, true);
          if (got) data[fullKey.slice(prefix.length)] = JSON.parse(got.value);
        } catch { /* clé illisible, on l'ignore plutôt que de faire échouer tout l'export */ }
      }
    }
  } catch { /* window.storage indisponible */ }
  // Complète avec le repli mémoire (utile hors environnement claude.ai)
  Object.keys(memoryStore).forEach(k => {
    if (k.startsWith(prefix) && !(k.slice(prefix.length) in data)) data[k.slice(prefix.length)] = memoryStore[k];
  });
  return { app: "SuiviIndiv", version: 1, exportedAt: new Date().toISOString(), teamId, teamName, data };
}

async function importAllData(teamId, backup) {
  const prefix = "team_" + teamId + ":";
  const entries = Object.entries(backup.data || {});
  for (const [relKey, value] of entries) {
    const fullKey = prefix + relKey;
    try {
      if (window.storage) await window.storage.set(fullKey, JSON.stringify(value), true);
      memoryStore[fullKey] = value;
    } catch { memoryStore[fullKey] = value; }
  }
  return entries.length;
}

// Une colonne de box score dont le nom contient "%" doit s'afficher en pourcentage — le
// fichier importé stocke parfois ça en fraction (0.1496), donc on la remultiplie par 100
// si besoin (heuristique : une vraie valeur en % dépasse rarement 1 en dessous de 100%).
function formatStatValue(label, value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "–";
  if (/%/.test(label)) {
    const pct = Math.abs(value) <= 1 ? value * 100 : value;
    return pct.toFixed(1) + "%";
  }
  return value.toFixed(1);
}

// Rang de ce joueur dans l'équipe pour une statistique de box score donnée (classement
// décroissant : suppose que "plus haut = mieux", vrai pour la majorité des stats de base).
function teamRank(byPlayer, statLabel, playerFirst) {
  const entries = Object.entries(byPlayer)
    .filter(([, b]) => b.averages[statLabel] !== undefined && b.averages[statLabel] !== null)
    .map(([name, b]) => ({ name, val: b.averages[statLabel] }))
    .sort((a, b) => b.val - a.val);
  const idx = entries.findIndex(e => e.name === playerFirst);
  if (idx === -1) return null;
  return { rank: idx + 1, total: entries.length };
}

function fileToResizedDataURL(file, maxDim = 320, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Lecture du fichier impossible."));
    reader.onload = () => {
      img.onerror = () => reject(new Error("Image invalide."));
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDim) { height = Math.round((height * maxDim) / width); width = maxDim; }
        else if (height > maxDim) { width = Math.round((width * maxDim) / height); height = maxDim; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

// Lit une capture d'écran d'un tableau de stats (photo/screenshot) via la vision de Claude
// et en extrait les stats de scouting dans le même schéma que le reste de l'appli.
async function extractScoutingFromImage(file) {
  const dataUrl = await fileToResizedDataURL(file, 1400, 0.9); // résolution plus haute pour la lecture, pas pour l'affichage
  const base64 = dataUrl.split(",")[1];
  const schemaHint = SCOUT_STAT_SCHEMA.map(s => `"${s.key}" (${s.label}${s.pct ? ", en %" : ""})`).join(", ");
  const prompt = `Cette image est une capture d'écran d'un tableau de statistiques d'une ou plusieurs équipes de basketball. ` +
    `Extrais les données dans EXACTEMENT ce format JSON, sans aucun texte avant ou après : ` +
    `{"teams":[{"name":"Team name as written in the image","stats":{"pts":18.5,...}}]}. ` +
    `Les clés de "stats" possibles sont : ${schemaHint}. N'inclus une clé QUE si tu es raisonnablement sûr de sa valeur lue dans l'image ; ` +
    `omets-la sinon plutôt que d'inventer un chiffre. Les pourcentages doivent être en valeur 0-100 (ex. 45.2, pas 0.452).`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1500,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
          { type: "text", text: prompt },
        ],
      }],
    }),
  });
  if (!response.ok) throw new Error("Erreur de lecture de l'image (API).");
  const data = await response.json();
  const text = (data.content || []).map(b => b.text || "").join("\n").trim();
  const cleaned = text.replace(/^```json\s*|```$/g, "").trim();
  let parsed;
  try { parsed = JSON.parse(cleaned); } catch { throw new Error("Reading the image did not return a usable format — try again with a clearer image."); }
  if (!parsed.teams || !Array.isArray(parsed.teams)) throw new Error("No team detected in the image.");
  return parsed.teams;
}

// ---------------------------------------------------------------------------
// Storage helpers (avec fallback mémoire si window.storage indisponible)
// ---------------------------------------------------------------------------

const memoryStore = {};
let TEAM_PREFIX = ""; // préfixe de toutes les clés une fois une équipe sélectionnée — isole totalement les données des deux équipes
function setActiveTeam(teamId) { TEAM_PREFIX = teamId ? "team_" + teamId + ":" : ""; }

let supabase = null;
const supabaseInit = (async () => {
  try {
    const url = import.meta.env.VITE_SUPABASE_URL;
    const key = import.meta.env.VITE_SUPABASE_ANON_KEY;
    if (url && key) {
      const { createClient } = await import("@supabase/supabase-js");
      supabase = createClient(url, key);
      console.log("[storage] Supabase connected:", url);
    } else {
      console.warn("[storage] Supabase env vars missing — falling back to local memory only.");
    }
  } catch (e) {
    console.error("[storage] Supabase init failed, falling back to local memory only:", e);
  }
})();

async function storeGet(key) {
  await supabaseInit;
  const k = TEAM_PREFIX + key;
  try {
    if (window.storage) {
      const r = await window.storage.get(k, true);
      return r ? JSON.parse(r.value) : null;
    }
    if (supabase) {
      const { data } = await supabase.from("app_storage").select("value").eq("key", k).maybeSingle();
      return data ? data.value : null;
    }
  } catch (e) { /* clé absente */ }
  return memoryStore[k] ?? null;
}
async function storeSet(key, value) {
  await supabaseInit;
  const k = TEAM_PREFIX + key;
  memoryStore[k] = value; // repli local immédiat, pour que l'UI reste cohérente même en cas d'échec réseau
  if (window.storage) {
    await window.storage.set(k, JSON.stringify(value), true);
    return;
  }
  if (supabase) {
    const { error } = await supabase.from("app_storage").upsert({ key: k, value });
    if (error) { console.error("[storage] Supabase set FAILED on", k, ":", error); throw error; }
    return;
  }
  console.warn("[storage] No persistent backend available — data for", k, "only exists in memory for this session.");
}
async function storeDelete(key) {
  await supabaseInit;
  const k = TEAM_PREFIX + key;
  try {
    if (window.storage) await window.storage.delete(k, true);
    else if (supabase) await supabase.from("app_storage").delete().eq("key", k);
  } catch (e) {}
  delete memoryStore[k];
}
async function storeList(prefix) {
  await supabaseInit;
  const p = TEAM_PREFIX + prefix;
  try {
    if (window.storage) {
      const r = await window.storage.list(p, true);
      return r ? r.keys : [];
    }
    if (supabase) {
      const { data } = await supabase.from("app_storage").select("key").like("key", p + "%");
      return (data || []).map(r => r.key);
    }
  } catch (e) {}
  return Object.keys(memoryStore).filter(k => k.startsWith(p));
}

// La session (qui est connecté) doit rester STRICTEMENT LOCALE à cet appareil/navigateur —
// jamais passer par le stockage partagé de l'équipe (Supabase), sinon un coach connecté sur
// un appareil se retrouve automatiquement connecté sur TOUS les autres appareils qui ouvrent
// le site, puisque ce stockage est synchronisé en temps réel entre tout le monde.
function sessionStorageKey() { return "hooptrack_session_" + TEAM_PREFIX; }
async function saveLocalSession(session) {
  try {
    if (typeof localStorage !== "undefined") { localStorage.setItem(sessionStorageKey(), JSON.stringify(session)); return; }
  } catch (e) {}
  // Repli pour l'artifact Claude, où localStorage est volontairement indisponible —
  // stockage "personnel" (shared=false), déjà scindé par utilisateur Claude.ai.
  try { if (window.storage) await window.storage.set(sessionStorageKey(), JSON.stringify(session), false); } catch (e) {}
}
async function loadLocalSession() {
  try {
    if (typeof localStorage !== "undefined") {
      const raw = localStorage.getItem(sessionStorageKey());
      return raw ? JSON.parse(raw) : null;
    }
  } catch (e) {}
  try {
    if (window.storage) { const r = await window.storage.get(sessionStorageKey(), false); return r ? JSON.parse(r.value) : null; }
  } catch (e) {}
  return null;
}
async function clearLocalSession() {
  try {
    if (typeof localStorage !== "undefined") { localStorage.removeItem(sessionStorageKey()); return; }
  } catch (e) {}
  try { if (window.storage) await window.storage.delete(sessionStorageKey(), false); } catch (e) {}
}

// Même principe que la session : quelle équipe est active doit rester STRICTEMENT LOCALE à
// cet appareil. Sinon, la dernière équipe choisie par n'importe qui se propage à tous les
// autres appareils qui ouvrent le site — bug constaté où des joueurs arrivaient directement
// sur une équipe sans jamais être passés par l'écran de choix.
const ACTIVE_TEAM_LOCAL_KEY = "hooptrack_active_team_id";
async function saveLocalActiveTeam(teamId) {
  try {
    if (typeof localStorage !== "undefined") { localStorage.setItem(ACTIVE_TEAM_LOCAL_KEY, teamId); return; }
  } catch (e) {}
  try { if (window.storage) await window.storage.set(ACTIVE_TEAM_LOCAL_KEY, teamId, false); } catch (e) {}
}
async function loadLocalActiveTeam() {
  try {
    if (typeof localStorage !== "undefined") { return localStorage.getItem(ACTIVE_TEAM_LOCAL_KEY); }
  } catch (e) {}
  try {
    if (window.storage) { const r = await window.storage.get(ACTIVE_TEAM_LOCAL_KEY, false); return r ? r.value : null; }
  } catch (e) {}
  return null;
}
async function clearLocalActiveTeam() {
  try {
    if (typeof localStorage !== "undefined") { localStorage.removeItem(ACTIVE_TEAM_LOCAL_KEY); return; }
  } catch (e) {}
  try { if (window.storage) await window.storage.delete(ACTIVE_TEAM_LOCAL_KEY, false); } catch (e) {}
}

// ---------------------------------------------------------------------------
// Parsing du fichier "Database"
// ---------------------------------------------------------------------------

// Vocabulaire fixe des tags du logiciel de coding (tout ce qui N'EST PAS un joueur).
// Tout en-tête de colonne qui ne figure pas dans cette liste, entre "button" et la
// première colonne vide, est considéré comme une colonne joueur — exactement le principe
// demandé : la détection des joueurs vient uniquement du fichier importé, pas d'une liste
// codée en dur.
// ---------------------------------------------------------------------------
// Scouting — comparaison d'équipes (stats brutes + avancées déjà calculatedes)
// ---------------------------------------------------------------------------

// Schéma calqué sur un vrai fichier "Stats Center" de championnat : ce sont déjà les
// stats avancées calculatedes (ORTG, DRTG, eFG%, TS%…), pas besoin de refaire les formules.
const SCOUT_STAT_SCHEMA = [
  { key: "mj", label: "Games played", group: "General" },
  { key: "pctv", label: "% Victoires", pct: true, group: "General" },
  { key: "min", label: "Minutes", group: "General" },
  { key: "pts", label: "Points", group: "General" },
  { key: "ptse", label: "Points allowed", lowerBetter: true, group: "General" },
  { key: "plusminus", label: "+/-", group: "General" },
  { key: "eff", label: "Evaluation", group: "General" },
  { key: "r2", label: "2pt made", group: "Shooting" },
  { key: "t2", label: "2pt attempted", group: "Shooting" },
  { key: "pct2", label: "% 2pts", pct: true, group: "Shooting" },
  { key: "r3", label: "3pt made", group: "Shooting" },
  { key: "t3", label: "3pt attempted", group: "Shooting" },
  { key: "pct3", label: "% 3pts", pct: true, group: "Shooting" },
  { key: "lfr", label: "FT made", group: "Shooting" },
  { key: "lft", label: "FT attempted", group: "Shooting" },
  { key: "pctlf", label: "% LF", pct: true, group: "Shooting" },
  { key: "efg", label: "eFG%", pct: true, group: "Shooting" },
  { key: "ts", label: "TS%", pct: true, group: "Shooting" },
  { key: "ftafga", label: "FTA/FGA", group: "Shooting" },
  { key: "ro", label: "Off. rebounds", group: "Rebounds & Ball" },
  { key: "rd", label: "Def. rebounds", group: "Rebounds & Ball" },
  { key: "rt", label: "Total rebounds", group: "Rebounds & Ball" },
  { key: "pctro", label: "% Off. rebounds", pct: true, group: "Rebounds & Ball" },
  { key: "pd", label: "Assists", group: "Rebounds & Ball" },
  { key: "bp", label: "Turnovers", lowerBetter: true, group: "Rebounds & Ball" },
  { key: "padbp", label: "Assists/Turnovers", group: "Rebounds & Ball" },
  { key: "pctbp", label: "% TOV", pct: true, lowerBetter: true, group: "Rebounds & Ball" },
  { key: "pctpad", label: "% Assists", pct: true, group: "Rebounds & Ball" },
  { key: "ct", label: "Blocks", group: "Defense" },
  { key: "int", label: "Steals", group: "Defense" },
  { key: "fte", label: "Fouls", lowerBetter: true, group: "Defense" },
  { key: "fo", label: "Offensive fouls", lowerBetter: true, group: "Defense" },
  { key: "poss", label: "Possessions", group: "Ratings" },
  { key: "ortg", label: "ORTG", group: "Ratings" },
  { key: "drtg", label: "DRTG", lowerBetter: true, group: "Ratings" },
  { key: "pct3tst", label: "3pt shot frequency (share of shots attempted)", pct: true, group: "Shooting" },
  // Métriques propres aux exports FIBA — gardées sous leur nom d'origine (pas retraduites)
  // pour éviter de leur donner un sens que je ne suis pas sûr de connaître avec certitude.
  { key: "fgAtt", label: "FG Att (FIBA)", group: "FIBA (advanced)" },
  { key: "fgMade", label: "FG Made (FIBA)", group: "FIBA (advanced)" },
  { key: "fgPct", label: "FG% (FIBA)", pct: true, group: "FIBA (advanced)" },
  { key: "ppp", label: "PPP (FIBA)", group: "FIBA (advanced)" },
  { key: "spi", label: "SPI (FIBA)", group: "FIBA (advanced)" },
  { key: "pps", label: "PPS (FIBA)", group: "FIBA (advanced)" },
  { key: "ssq", label: "SSQ (FIBA)", group: "FIBA (advanced)" },
  { key: "ssm", label: "SSM (FIBA)", group: "FIBA (advanced)" },
  { key: "pctft_fiba", label: "%FT (FIBA — meaning to confirm)", pct: true, group: "FIBA (advanced)" },
  { key: "pctsf", label: "%SF (FIBA)", pct: true, group: "FIBA (advanced)" },
  { key: "scorepct", label: "Score% (FIBA)", pct: true, group: "FIBA (advanced)" },
  // Issues du fichier FIBA "Defensive" — ce que l'équipe CONCÈDE ou FORCE chez l'adversaire.
  { key: "dPct2", label: "% 2pt allowed", pct: true, lowerBetter: true, group: "Defense (opponent)" },
  { key: "dPct3", label: "% 3pt allowed", pct: true, lowerBetter: true, group: "Defense (opponent)" },
  { key: "dEfg", label: "eFG% allowed", pct: true, lowerBetter: true, group: "Defense (opponent)" },
  { key: "dPctbp", label: "% Turnovers forced", pct: true, group: "Defense (opponent)" },
  { key: "dFtafga", label: "FTA/FGA allowed", lowerBetter: true, group: "Defense (opponent)" },
];
const SCOUT_KEY_BY_LABEL = {};
SCOUT_STAT_SCHEMA.forEach(s => { SCOUT_KEY_BY_LABEL[normTag(s.label)] = s.key; });

// En-têtes possibles telles qu'elles apparaissent dans un vrai fichier "Stats Center".
const SCOUT_HEADER_ALIASES = {
  mj: ["mj"], pctv: ["%v"], min: ["min"], pts: ["pts"], r2: ["2r"], t2: ["2t"], pct2: ["%2p"],
  r3: ["3r"], t3: ["3t"], pct3: ["%3p"], lfr: ["lfr"], lft: ["lft"], pctlf: ["%lf"],
  ro: ["ro"], rd: ["rd"], rt: ["rt"], pd: ["pd"], ct: ["ct"], int: ["int"], bp: ["bp"],
  fte: ["fte"], fo: ["fo"], plusminus: ["+/-"], eff: ["eff"], ptse: ["ptse", "pts e"],
  poss: ["poss"], ortg: ["ortg"], drtg: ["drtg"], padbp: ["pad/bp"], pctbp: ["%bp"],
  pctro: ["%ro"], pct3tst: ["%3ptst"], pctpad: ["%pad"], ftafga: ["fta/fga"], efg: ["efg%"], ts: ["ts%"],
};

function parseScoutingExcelFile(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames.find(n => /stat.?cent/i.test(n)) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: "" });
  const headerRowIdx = rows.findIndex(r => r && String(r[0] ?? "").trim().toLowerCase() === "club");
  if (headerRowIdx === -1) throw new Error("'Stats Center' sheet not found ('Club' header not detected).");
  const headerRow = rows[headerRowIdx].map(h => String(h ?? "").trim());

  const colToKey = {};
  headerRow.forEach((h, i) => {
    if (!h) return;
    const norm = normTag(h);
    for (const [key, aliases] of Object.entries(SCOUT_HEADER_ALIASES)) {
      if (aliases.some(a => normTag(a) === norm)) { colToKey[i] = key; break; }
    }
  });

  const teams = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row && row[0] ? String(row[0]).trim() : "";
    if (!name || name.toLowerCase() === "club") continue;
    const stats = {};
    Object.entries(colToKey).forEach(([idx, key]) => {
      const v = row[Number(idx)];
      if (v instanceof Date) return; // artefact Excel (ex. colonne V-D mal formatée), on ignore
      if (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) {
        const schema = SCOUT_STAT_SCHEMA.find(s => s.key === key);
        stats[key] = schema && schema.pct && Math.abs(Number(v)) <= 1 ? Number(v) * 100 : Number(v);
      }
    });
    if (Object.keys(stats).length) teams.push({ name, stats });
  }
  return { sheetName, teams };
}

// En-têtes FIBA -> clés du schéma. Les colonnes dont le sens n'est pas garanti restent
// sous leur nom d'origine (groupe "FIBA (advanced)") plutôt que d'être réinterprétées.
const FIBA_HEADER_MAP_OFF = {
  "gp": "mj", "poss": "poss", "pts": "pts", "ppp": "ppp",
  "fg att": "fgAtt", "fg made": "fgMade", "fg%": "fgPct", "efg%": "efg",
  "to%": "pctbp", "fta/fga": "ftafga", "spi": "spi", "pps": "pps", "ssq": "ssq", "ssm": "ssm",
  "%ft": "pctft_fiba", "%sf": "pctsf", "score%": "scorepct",
  "2 fg att": "t2", "2 fg made": "r2", "2 fg%": "pct2",
  "3fg att": "t3", "3 fg made": "r3", "3 fg%": "pct3", "3pa/fga": "pct3tst",
};

// Fichier FIBA "Defensive" : mêmes colonnes mais ce qu'on CONCÈDE à l'adversaire. La
// finesse demandée : "PPP" y devient le DRTG (points encaissés par possession, x100 pour
// être sur la même échelle que l'ORTG). "Pts" y devient les points encaissés (ptse).
const FIBA_HEADER_MAP_DEF = {
  "gp": "mj", "poss": "poss", "pts": "ptse", "ppp": "drtg",
  "fg%": "dEfg" /* fallback si eFG% absent */, "efg%": "dEfg",
  "to%": "dPctbp", "fta/fga": "dFtafga",
  "2 fg%": "dPct2", "3 fg%": "dPct3",
};

function parseFibaLeaderboardFile(arrayBuffer, side = "offense") {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: "" });
  const headerRowIdx = rows.findIndex(r => r && r.some(c => String(c ?? "").trim().toLowerCase() === "team"));
  if (headerRowIdx === -1) throw new Error("Format non reconnu (colonne 'Team' introuvable).");
  const headerRow = rows[headerRowIdx].map(h => String(h ?? "").trim());
  const teamColIdx = headerRow.findIndex(h => h.toLowerCase() === "team");
  const map = side === "defense" ? FIBA_HEADER_MAP_DEF : FIBA_HEADER_MAP_OFF;

  const colToKey = {};
  headerRow.forEach((h, i) => {
    const key = map[h.toLowerCase().trim()];
    if (key) colToKey[i] = key;
  });
  if (!Object.keys(colToKey).length) throw new Error("Aucune colonne FIBA reconnue dans ce fichier.");

  const teams = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row && row[teamColIdx] ? String(row[teamColIdx]).trim() : "";
    if (!name) continue;
    const stats = {};
    Object.entries(colToKey).forEach(([idx, key]) => {
      const v = row[Number(idx)];
      if (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) {
        const schema = SCOUT_STAT_SCHEMA.find(s => s.key === key);
        // Les valeurs FIBA sont en fraction (0-1) pour les %, on les remet en 0-100.
        // Le DRTG dérivé de "PPP" (points par possession, ~0.8-1.1) suit la même mise à
        // l'échelle x100 que l'ORTG, même s'il n'est pas marqué comme un "%" au sens strict.
        stats[key] = (schema && schema.pct) || key === "drtg" ? Number(v) * 100 : Number(v);
      }
    });
    if (side === "offense" && stats.pts !== undefined && stats.poss) {
      // ORTG calculable directement (Points/Possessions), même si FIBA ne le fournit pas
      // dans l'export "Offensive" — DRTG lui vient du fichier "Defensive" (PPP ci-dessus).
      stats.ortg = (100 * stats.pts) / stats.poss;
    }
    if (Object.keys(stats).length) teams.push({ name, stats });
  }
  return { sheetName, teams };
}

// Catégories de colonnes du fichier de coding — modifiables par le coach (onglet
// Settings) sans toucher au code, pour absorber les évolutions du logiciel de coding.
// Ces valeurs par défaut reproduisent exactement la catégorisation d'origine.
const DEFAULT_TAG_CATEGORIES = {
  "Player": DEFAULT_ROSTER.map(p => p.name),
  "Playtypes": ["1v1", "Backdoor", "Backpick", "Cut", "Fastbreak", "Flarescreen", "Hand Off", "Offscreen", "PnP", "PnR", "Postup", "Rebound", "Spain", "StepUp", "StepUp PnP", "Transition", "Zone"],
  "Plays": ["In", "Thumb", "Away", "Chase", "Play A", "Play B", "Play C", "Play D", "Play E", "BOB", "SOB", "S"],
  "Shot selection": ["Contested", "Open"],
  "Defensive mistakes": ["Ball", "First Pass", "Last", "OBS", "RBD", "RNJ", "Repli", "Safety", "Screener"],
  "Screen defense": ["Ice", "Protect", "Reject", "Step", "Switch", "Under"],
  "Spacing": ["0 in Front", "1 in Front", "2 in Front", "3 in Front", "3 Baseline"],
  "Shot zone": ["Weak", "Corner 3", "Front 3", "Midrange", "Paint"],
  "Defense type": ["Man to man", "Zone", "Match-up zone", "Press", "Box and one"],
  "Results & misc.": [
    "0FT+", "1FT+", "1FTA", "2FT+", "2FTA", "2PT+", "2PT-", "3FT+", "3FTA", "3PT+", "3PT-",
    "AND 1", "FOUL", "FT FOUL", "OREB", "OUT OF BOUNDS", "TOV",
  ],
};

// État mutable au niveau module (même schéma que TEAM_PREFIX) — chargé une fois par
// équipe active, modifié en direct depuis l'onglet Settings, utilisé partout ensuite
// sans avoir à faire redescendre la config par les props dans tout l'arbre.
let TAG_CATEGORIES = null;
function currentTagCategories() { return TAG_CATEGORIES || DEFAULT_TAG_CATEGORIES; }
function categoryTags(name, cats) { return (cats || currentTagCategories())[name] || []; }
function allKnownTagsSet(cats) {
  return new Set(Object.entries(cats || currentTagCategories()).filter(([name]) => name !== "Player").flatMap(([, tags]) => tags).map(normTag));
}
function knownPlayersSet(cats) {
  return new Set(categoryTags("Player", cats).map(normTag));
}
// Certaines équipes ont sauvegardé leurs catégories AVANT la traduction en anglais — cette
// table migre silencieusement les anciens noms français vers les nouveaux à chaque chargement,
// pour ne pas laisser de vieilles données françaises persister indéfiniment.
const CATEGORY_NAME_MIGRATIONS = {
  "Joueur": "Player",
  "Défense d'écran": "Screen defense",
  "Défenses d'écran": "Screen defense",
  "Zone de tir": "Shot zone",
  "Erreurs défensives": "Defensive mistakes",
  "Sélection de tir": "Shot selection",
  "Résultats & divers": "Results & misc.",
};
async function loadTagCategories() {
  const stored = await storeGet("tag_categories");
  if (stored) {
    let migrated = false;
    const next = {};
    for (const [key, val] of Object.entries(stored)) {
      const newKey = CATEGORY_NAME_MIGRATIONS[key];
      if (newKey && !stored[newKey]) { next[newKey] = val; migrated = true; }
      else next[key] = val;
    }
    TAG_CATEGORIES = next;
    if (migrated) await storeSet("tag_categories", next);
  } else {
    TAG_CATEGORIES = JSON.parse(JSON.stringify(DEFAULT_TAG_CATEGORIES));
    await storeSet("tag_categories", TAG_CATEGORIES);
  }
}
async function saveTagCategories(cats) {
  TAG_CATEGORIES = cats;
  await storeSet("tag_categories", cats);
}

// BUG RÉEL CORRIGÉ : la catégorie "Player" (Settings) est une liste SÉPARÉE du roster, qui
// pouvait se désynchroniser silencieusement (joueur renommé, nouveau joueur ajouté, sans
// jamais rouvrir Settings pour la resynchroniser). Un joueur pourtant bien enregistré était
// alors traité comme "non reconnu" à l'import d'un fichier de coding, et son nom brut de
// colonne devenait sa nouvelle identité au lieu de son vrai nom — donnant l'impression qu'un
// nouveau joueur avait été "créé" (doublon). On resynchronise donc TOUJOURS cette liste
// depuis le roster actuel juste avant de lire un fichier de coding, plutôt que de compter sur
// une synchronisation manuelle que le coach pourrait oublier.
async function syncPlayerCategoryFromRoster(roster) {
  const cats = currentTagCategories();
  const rosterNames = roster.map(p => p.name);
  const existing = cats["Player"] || [];
  const merged = Array.from(new Set([...rosterNames, ...existing]));
  if (merged.length === existing.length && merged.every((n, i) => n === existing[i])) return; // déjà à jour
  await saveTagCategories({ ...cats, Player: merged });
}

// Deuxième jeu de catégories, totalement indépendant du premier — utilisé uniquement par
// l'onglet Observation (Scouting), pour que ses colonnes/tags puissent différer de ceux
// du fichier de coding utilisé dans Import Match.
let OBSERVATION_TAG_CATEGORIES = null;
function currentObservationTagCategories() { return OBSERVATION_TAG_CATEGORIES || DEFAULT_TAG_CATEGORIES; }
async function loadObservationTagCategories() {
  const stored = await storeGet("observation_tag_categories");
  OBSERVATION_TAG_CATEGORIES = stored || JSON.parse(JSON.stringify(DEFAULT_TAG_CATEGORIES));
  if (!stored) await storeSet("observation_tag_categories", OBSERVATION_TAG_CATEGORIES);
}
async function saveObservationTagCategories(cats) {
  OBSERVATION_TAG_CATEGORIES = cats;
  await storeSet("observation_tag_categories", cats);
}

// Style de graphique choisi pour chaque catégorie PERSONNALISÉE (celles ajoutées dans
// Settings, au-delà des catégories intégrées comme Playtypes qui ont déjà un graphique
// dédié en dur) : "simple" = donut de comparaison entre les éléments eux-mêmes (comme
// Shooting Selection), "detailed" = liste avec fréquence/PPPP/Open% (comme Plays).
let CATEGORY_CHART_STYLES = null;
function currentCategoryChartStyles() { return CATEGORY_CHART_STYLES || {}; }
function chartStyleFor(categoryName) { return currentCategoryChartStyles()[categoryName] || "detailed"; }
async function loadCategoryChartStyles() {
  CATEGORY_CHART_STYLES = (await storeGet("category_chart_styles")) || {};
}
async function saveCategoryChartStyles(styles) {
  CATEGORY_CHART_STYLES = styles;
  await storeSet("category_chart_styles", styles);
}

// unknownColumnDefault : que faire d'une colonne qui n'est reconnue ni comme tag connu, ni
// comme joueur déjà confirmé dans les paramètres.
// - "player" (par défaut, utilisé par Import Match) : la traite comme un joueur potentiel à
//   confirmer — utile quand le fichier vient de NOTRE équipe et qu'un joueur a pu être ajouté
//   sans être encore synchronisé dans "Player" (Settings).
// - "tag" (utilisé par Observation) : la traite comme un tag supplémentaire au lieu d'un faux
//   joueur — BUG RÉEL CORRIGÉ : un fichier de scouting d'équipe adverse n'a souvent AUCUN nom
//   de joueur individuel (juste des tags d'équipe comme systèmes de jeu ou types de défense) ;
//   avec l'ancien comportement, chaque tag non reconnu (ex. "UCLA", "Corner Flare", "Follow")
//   devenait un faux "joueur", donnant des actions du type player:"UCLA" — n'ayant aucun sens.
function parseMatchFile(arrayBuffer, cats, unknownColumnDefault = "player") {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("base") || n.toLowerCase().includes("data")) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: 0 });
  if (!rows.length) throw new Error("Feuille vide ou introuvable.");

  // La ligne d'en-tête n'est pas forcément la 1ère ligne : certains exports du logiciel de
  // coding mettent le nom du match en ligne 1 et les en-têtes (category, button, …) en ligne 2.
  // On cherche donc la première ligne dont la colonne A vaut "category".
  let headerRowIdx = rows.findIndex(r => r && String(r[0] ?? "").trim().toLowerCase() === "category");
  if (headerRowIdx === -1) headerRowIdx = 0;
  const headerRow = (rows[headerRowIdx] || []).map(h => (h === undefined || h === null ? "" : String(h).trim()));

  // La frontière = première colonne vide après la colonne B (button), qui correspond
  // à la première colonne rouge du fichier d'origine : au-delà, ce sont des formules
  // ajoutées manuellement, pas les données brutes du logiciel de coding.
  let boundary = headerRow.length;
  for (let i = 2; i < headerRow.length; i++) {
    if (!headerRow[i]) { boundary = i; break; }
  }
  const cols = headerRow.slice(0, boundary);
  if (cols.length < 3) throw new Error("Unable to detect data columns ('category' / 'button' headers not found).");

  const playerColIdx = [];
  const tagColIdx = [];
  const knownPlayers = knownPlayersSet(cats);
  cols.forEach((c, i) => {
    if (i <= 1 || !c) return;
    const trimmed = String(c).trim();
    if (knownPlayers.has(normTag(c))) { playerColIdx.push(i); return; } // in the "Player" list → confirmed
    if (allKnownTagsSet(cats).has(normTag(c))) { tagColIdx.push(i); return; }
    if (/^\d+$/.test(trimmed)) return; // en-tête purement numérique (ex. "0") = ligne d'équipe, pas un joueur
    if (unknownColumnDefault === "tag") { tagColIdx.push(i); return; }
    playerColIdx.push(i); // neither a known tag nor in the "Player" list → treated as a player by default, to confirm
  });
  const detectedPlayers = playerColIdx.map(i => cols[i]);
  const unconfirmedPlayers = detectedPlayers.filter(p => !knownPlayers.has(normTag(p)));

  // Les données commencent juste après la ligne d'en-tête. On filtre les lignes vides.
  const dataRows = rows.slice(headerRowIdx + 1).filter(r => Array.isArray(r) && r.some(v => v !== 0 && v !== "" && v !== undefined));

  const plays = dataRows.flatMap(r => {
    const category = r[0] ?? "";
    const button = r[1] ?? "";
    const tags = {};
    for (const i of tagColIdx) {
      const v = r[i];
      if (v !== 0 && v !== "" && v !== undefined && v !== null) tags[cols[i]] = v;
    }
    // Une ligne peut concerner plusieurs joueurs à la fois (ex. porteur de balle + écran
    // sur la même possession) : on crée une action par joueur marqué "1", pour ne perdre
    // aucune donnée — c'était la cause des actions manquantes constatées.
    // Si AUCUN joueur n'est marqué (ex. scouting d'équipe adverse sans suivi individuel), on
    // garde quand même l'action au niveau de l'équipe plutôt que de la perdre silencieusement.
    const flaggedPlayers = playerColIdx.filter(i => Number(r[i]) === 1).map(i => cols[i]);
    if (flaggedPlayers.length === 0) return [{ category: String(category), button: String(button), player: null, tags }];
    return flaggedPlayers.map(player => ({ category: String(category), button: String(button), player, tags }));
  });

  return { sheetName, columnsDetected: cols.length, boundaryColumn: boundary, totalRows: dataRows.length, playsWithPlayer: plays.length, detectedPlayers, unconfirmedPlayers, plays };
}

function playPoints(tags) {
  let pts = 0;
  for (const [name, val] of Object.entries(tags)) {
    if (!val) continue;
    const m2 = String(name).match(/^(\d)\s*PT\+$/i);
    if (m2) { pts += parseInt(m2[1], 10); continue; }
    if (/FT\+$/i.test(name)) pts += 1;
  }
  return pts;
}
function isMiss(tags) {
  return Object.keys(tags).some(k => /^(\d)\s*PT-$/i.test(k) && tags[k]);
}
function isMake(tags) {
  return Object.keys(tags).some(k => /^(\d)\s*PT\+$/i.test(k) && tags[k]);
}
function isTOV(tags) {
  return !!tags["TOV"];
}
function isOffense(p) { return /offense/i.test(p.button); }
function isDefense(p) { return /defense/i.test(p.button); }

const EXCLUDED_FROM_BREAKDOWN = new Set([
  "0FT+","1FT+","1FTA","2FT+","2FTA","0FT-","2PT+","2PT-","3FT+","3FTA","3PT+","3PT-",
  "AND 1","FOUL","FT FOUL","OREB","OUT OF BOUNDS","TOV"
]);

// Playtypes (comment on joue le ballon) et Plays (appels de jeu / entrées) — labels canoniques
// tolérants aux variations d'écriture du fichier source (espaces, casse, PnR/PNR…).
// Ce sont maintenant des fonctions (pas des constantes figées) pour refléter en direct les
// catégories modifiées dans l'onglet Settings.
function PLAYTYPES_LIST(cats) { return categoryTags("Playtypes", cats); }
function PLAYS_LIST(cats) { return categoryTags("Plays", cats); }
function DEFENSIVE_MISTAKES_LIST(cats) { return categoryTags("Defensive mistakes", cats); }
// Type de couverture défensive sur écran (comment on défend un pick, ou comment
// l'adversaire défend les nôtres) — seuls les tags de couverture proprement dits.
function SCREEN_DEFENSE_LIST(cats) { return categoryTags("Screen defense", cats); }
// Spacing joué au moment d'un écran (nombre de joueurs "in front"/en retrait) — catégorie
// distincte de la couverture défensive elle-même.
function SPACING_LIST(cats) { return categoryTags("Spacing", cats); }

// Reproduit la logique du Dashboard original : top N tags par fréquence, avec un
// optional "Other" bucket grouping the rest (used for the "Plays" donut).
function topBucket(plays, labels, n, withOther) {
  const items = groupBreakdown(plays, labels).sort((a, b) => b.count - a.count).slice(0, n);
  const top = items.map(i => ({ name: i.label, value: i.count, freq: i.freq, pppp: i.pppp, open: i.open }));
  if (withOther) {
    const used = top.reduce((s, i) => s + i.value, 0);
    const other = plays.length - used;
    if (other > 0) top.push({ name: "Autres", value: other, freq: plays.length ? (100 * other) / plays.length : 0, pppp: null, open: null });
  }
  return top;
}

// Comme groupBreakdown, mais la fréquence de chaque item est calculatede par rapport au TOTAL
// DES ACTIONS DE CETTE CATÉGORIE (ex. total des "Plays" tagués), pas par rapport à
// l'ensemble des actions offensives/défensives. C'est ce que fait le fichier d'origine pour
// les "Plays" : 4 actions "In" sur 11 actions taguées play = 36%, pas 4 sur 28.
function categoryBreakdown(plays, labels) {
  const items = groupBreakdown(plays, labels);
  const categoryTotal = items.reduce((s, g) => s + g.count, 0);
  return items.map(g => ({ ...g, freq: categoryTotal ? (100 * g.count) / categoryTotal : 0 }));
}

function topCategoryBucket(plays, labels, n) {
  return categoryBreakdown(plays, labels)
    .sort((a, b) => b.count - a.count)
    .slice(0, n)
    .map(i => ({ name: i.label, value: i.count, freq: i.freq, pppp: i.pppp, open: i.open }));
}

// Répartition "Shooting Selection" : Ouvert vs Contesté/Perte de balle — mêmes catégories
// que le Dashboard original (bucket Contested regroupe aussi les pertes de balle).
function isShotAttempt(tags) {
  return Object.keys(tags).some(k => /^\d\s*PT[+-]$/i.test(k) && tags[k]);
}

function shootingSelection(plays, cats) {
  // Retracé directement depuis les formules du classeur original (cellule B30 d'une feuille
  // joueur) : le total de référence = la somme des actions taguées avec un playtype — pas
  // seulement celles qui ont un tag de résultat de tir explicite. C'est le même total qui
  // sert de base au "Ouvert" par playtype.
  const relevant = plays.filter(p => PLAYTYPES_LIST(cats).some(label => tagIsSet(p.tags, label)));
  const open = relevant.filter(p => tagIsSet(p.tags, "Open") || tagIsSet(p.tags, "Ouvert")).length;
  const other = relevant.length - open;
  return [
    { name: "Open", value: open, color: TEAL },
    { name: "Contested / Turnover", value: other, color: RED },
  ].filter(d => d.value > 0);
}

function normTag(s) { return String(s).toLowerCase().replace(/[\s\-_]/g, ""); }

function tagIsSet(tags, canonicalLabel) {
  const norm = normTag(canonicalLabel);
  for (const [k, v] of Object.entries(tags)) { if (normTag(k) === norm && v) return true; }
  return false;
}

function openPct(matching) {
  if (!matching.length) return null;
  // La formule d'origine (vérifiée dans le classeur) divise le nombre d'actions "Ouvert"
  // par le TOTAL des actions de ce tag — pas seulement celles marquées Ouvert ou Contesté.
  // Diviser par Open+Contesté (comme je le faisais avant) donnait un % trop élevé dès
  // qu'une partie des actions n'étaient taguées ni Ouvert ni Contesté.
  const openCount = matching.filter(p => tagIsSet(p.tags, "Open") || tagIsSet(p.tags, "Ouvert")).length;
  return (100 * openCount) / matching.length;
}

function groupBreakdown(plays, labels) {
  return labels.map(label => {
    const matching = plays.filter(p => tagIsSet(p.tags, label));
    const pts = matching.reduce((s, p) => s + playPoints(p.tags), 0);
    return {
      label,
      count: matching.length,
      freq: plays.length ? (100 * matching.length) / plays.length : 0,
      pppp: matching.length ? pts / matching.length : 0,
      open: openPct(matching),
    };
  }).filter(g => g.count > 0);
}

// ---------------------------------------------------------------------------
// Parsing d'un fichier de stats de match complètes (box score)
// ---------------------------------------------------------------------------

function parseBoxScoreFile(arrayBuffer, roster, teamName) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  // On cherche la feuille la plus plausible : celle qui contient le plus de noms de joueurs du roster.
  const rosterFirsts = roster.map(p => p.first.toLowerCase());
  let bestSheet = wb.SheetNames[0], bestScore = -1, bestRows = null;
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" });
    let score = 0;
    rows.slice(0, 60).forEach(r => (r || []).forEach(c => {
      const s = String(c || "").toLowerCase();
      if (rosterFirsts.some(f => s.includes(f))) score++;
    }));
    if (score > bestScore) { bestScore = score; bestSheet = name; bestRows = rows; }
  }
  const rows = bestRows || [];
  if (!rows.length) throw new Error("Feuille vide ou introuvable.");

  // Beaucoup de box scores ont un EN-TÊTE SUR DEUX LIGNES : une ligne de groupe fusionnée
  // ("Tirs Tot.", "LF", "Reb"…) followed by a row of sub-categories ("Made", "Missed",
  // "Tot.", "%", "Def", "Off"…). Ne lire que la 1ère ligne fait pointer sur la mauvaise
  // sous-colonne (ex. "Made" au lieu de "Tot." pour les tirs tentés). On détecte ce cas et
  // on reconstruit un en-tête combiné "Group Subcategory".
  const row0 = rows[0] || [];
  const row1 = rows[1] || [];
  // Les fichiers réels ne mettent pas toujours un sous-en-tête "propre" ("Réussi" tout seul) :
  // certains répètent le nom du groupe et des tabulations dans la même cellule
  // ("2pts\t\t\tRéussi", "Reb\t\tDef"...). On ne regarde donc que le DERNIER segment utile
  // (après avoir découpé sur tabulations/espaces multiples), qui est le vrai sous-libellé.
  const subHeaderTokens = /^(r[ée]ussis?|manqu[ée]s?|tot\.?|%|def|off|made|missed|att|attempts?)$/i;
  function lastSubToken(cell) {
    const parts = String(cell ?? "").split(/\t+/).map(s => s.trim()).filter(Boolean);
    return parts.length ? parts[parts.length - 1] : "";
  }
  const row1LooksLikeSubHeader = row1.filter(c => c !== "" && c !== undefined && c !== null).length >= 3
    && row1.filter(c => c !== "" && subHeaderTokens.test(lastSubToken(c))).length >= 2;

  let header, dataStartRow;
  if (row1LooksLikeSubHeader) {
    // Propage le libellé de groupe vers la droite tant qu'on ne rencontre pas un nouveau libellé.
    const filledGroup = [];
    let last = "";
    for (let i = 0; i < Math.max(row0.length, row1.length); i++) {
      const g = String(row0[i] ?? "").trim();
      if (g) last = g;
      filledGroup.push(last);
    }
    // Certaines colonnes sont toujours des stats "simples" (jamais de sous-catégorie
    // Made/Missed/Tot./% qui leur soit propre) — on ne leur accole jamais de sous-en-tête,
    // même si une cellule voisine de la ligne 2 contient un mot qui y ressemble par erreur.
    const NEVER_SPLIT = new Set(["+/-", "ct", "int", "ev", "bp", "pad", "pts", "mj", "gp", "min"]);
    header = filledGroup.map((g, i) => {
      if (NEVER_SPLIT.has(g.trim().toLowerCase())) return g;
      const sub = lastSubToken(row1[i]);
      return sub ? `${g} ${sub}`.trim() : g;
    });
    dataStartRow = 2;
  } else {
    header = row0.map(h => String(h ?? "").trim());
    dataStartRow = 1;
  }

  const statCols = header.map((h, i) => ({ idx: i, label: h })).filter(c => c.idx > 0 && c.label);
  const teamNameLower = teamName ? teamName.trim().toLowerCase() : null;

  const parsedRows = [];
  const unmatchedRows = [];
  const candidateTeamRows = []; // toutes les lignes non-joueur avec des stats (peut inclure l'adversaire)
  for (let r = dataStartRow; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row[0] === undefined || row[0] === "") continue;
    const raw = String(row[0]).trim();
    const rawLower = raw.toLowerCase();
    const match = roster.find(p => rawLower.includes(p.first.toLowerCase()) || rawLower === p.name.toLowerCase());
    if (!match) {
      if (/^\d+$/.test(raw)) continue; // ligne purement numérique, ignorée
      const stats = {};
      statCols.forEach(c => {
        const v = row[c.idx];
        if (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) stats[c.label] = Number(v);
      });
      if (Object.keys(stats).length) candidateTeamRows.push({ label: raw, stats });
      else if (!/^(total|team|équipe|equipe)s?$/i.test(raw)) unmatchedRows.push(raw);
      continue;
    }
    const stats = {};
    statCols.forEach(c => {
      const v = row[c.idx];
      if (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) stats[c.label] = Number(v);
    });
    if (Object.keys(stats).length) parsedRows.push({ player: match.name, playerFull: match.name, stats });
  }

  // On identifie LA ligne de totaux d'équipe par son nom exact (renseigné par le coach),
  // pas par une supposition d'ordre — un fichier qui liste aussi les stats de l'adversaire
  // aurait sinon pu faire capturer la mauvaise ligne (celle de l'adversaire).
  let teamRow = null;
  let teamRowConfirmed = false;
  if (teamNameLower) {
    const found = candidateTeamRows.find(c => c.label.toLowerCase().includes(teamNameLower));
    if (found) { teamRow = found; teamRowConfirmed = true; }
  }
  if (!teamRow && candidateTeamRows.length === 1) { teamRow = candidateTeamRows[0]; teamRowConfirmed = false; }
  // Les autres lignes candidates non retenues comme ligne d'équipe repassent en "non reconnues".
  candidateTeamRows.forEach(c => { if (c !== teamRow) unmatchedRows.push(c.label); });

  const statLabels = Array.from(new Set(parsedRows.flatMap(r => Object.keys(r.stats))));
  const matchedPlayers = Array.from(new Set(parsedRows.map(r => r.player)));
  return { sheetName: bestSheet, statLabels, matchedCount: parsedRows.length, matchedPlayers, unmatchedRows, teamRow, teamRowConfirmed, rows: parsedRows };
}

// ---------------------------------------------------------------------------
// Petits composants UI
// ---------------------------------------------------------------------------

const CHART_COLORS = [AMBER, TEAL, "#7C9CF2", "#C97BE0", "#E15A4E", "#5CC98A", "#E0B85C", "#5CB3E0"];

function DonutChart({ data, size = 150, unit = "actions" }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  const cx = size / 2, cy = size / 2;
  const outerR = size * 0.48, innerR = size * 0.32;

  // SVG avec attribut "fill" explicite sur chaque tranche — mieux supporté à la fois par
  // les navigateurs ET par html2canvas (utilisé pour l'export PDF) qu'un dégradé CSS
  // conic-gradient, qui n'est pas fidèlement restitué par la capture PDF.
  function arcPath(startAngle, endAngle, r) {
    const x0 = cx + r * Math.cos(startAngle), y0 = cy + r * Math.sin(startAngle);
    const x1 = cx + r * Math.cos(endAngle), y1 = cy + r * Math.sin(endAngle);
    const large = endAngle - startAngle > Math.PI ? 1 : 0;
    return { x0, y0, x1, y1, large };
  }
  function donutSlicePath(startAngle, endAngle) {
    const o0 = arcPath(startAngle, endAngle, outerR), i0 = arcPath(startAngle, endAngle, innerR);
    return [
      `M ${o0.x0} ${o0.y0}`,
      `A ${outerR} ${outerR} 0 ${o0.large} 1 ${o0.x1} ${o0.y1}`,
      `L ${i0.x1} ${i0.y1}`,
      `A ${innerR} ${innerR} 0 ${i0.large} 0 ${i0.x0} ${i0.y0}`,
      "Z",
    ].join(" ");
  }

  let cursor = -Math.PI / 2;
  const gapRad = total > 0 ? 0.02 : 0;
  const slices = total > 0 ? data.filter(d => d.value > 0).map((d, i) => {
    const angle = (d.value / total) * (2 * Math.PI);
    const start = cursor + gapRad / 2, end = cursor + angle - gapRad / 2;
    cursor += angle;
    return { path: end > start ? donutSlicePath(start, end) : null, color: d.color || CHART_COLORS[i % CHART_COLORS.length], key: i };
  }) : [];

  return (
    <div style={{ position: "relative", width: size, height: size }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        {slices.map(s => s.path && <path key={s.key} d={s.path} fill={s.color} className="keep-color" />)}
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
        <div className="keep-color" style={{ fontFamily: "ui-monospace, monospace", fontSize: 20, fontWeight: 700, color: PAPER }}>{total}</div>
        <div style={{ fontSize: 9, color: "#5C6470", textTransform: "uppercase" }}>{unit}</div>
      </div>
    </div>
  );
}

function DonutLegend({ data }) {
  const total = data.reduce((s, d) => s + d.value, 0) || 1;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      {data.map((d, i) => {
        // Si un "freq" (% du total de possessions) est déjà calculated en amont, on l'utilise
        // pour rester cohérent avec les autres tableaux — sinon on retombe sur la part du
        // total affiché dans CE donut (cas Shooting Selection / Defensive mistakes, qui
        // forment une partition complète à 100%).
        const pct = d.freq !== undefined && d.freq !== null ? d.freq : (100 * d.value) / total;
        return (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
            <span className="keep-color" style={{ width: 9, height: 9, borderRadius: 3, background: d.color || CHART_COLORS[i % CHART_COLORS.length], flexShrink: 0 }} />
            <span style={{ color: "#8B93A1", flex: 1 }}>{d.name}</span>
            <span style={{ fontFamily: "ui-monospace, monospace", color: PAPER }}>{d.value}</span>
            <span style={{ fontFamily: "ui-monospace, monospace", color: "#5C6470", width: 38, textAlign: "right" }}>{pct.toFixed(0)}%</span>
          </div>
        );
      })}
    </div>
  );
}

function BreakdownChart({ title, data }) {
  if (!data.length) return null;
  const chartData = data.map(d => ({ name: d.label, "Frequency %": Number(d.freq.toFixed(1)), "PPPP": Number(d.pppp.toFixed(2)) }));
  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "18px 10px 8px 0" }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, padding: "0 18px 12px" }}>{title}</div>
      <ResponsiveContainer width="100%" height={Math.max(180, chartData.length * 34)}>
        <ComposedChart data={chartData} layout="vertical" margin={{ top: 0, right: 24, left: 0, bottom: 0 }}>
          <CartesianGrid stroke={LINE} horizontal={false} />
          <XAxis type="number" xAxisId="freq" orientation="top" tick={{ fill: "#5C6470", fontSize: 11 }} tickFormatter={v => v + "%"} axisLine={{ stroke: LINE }} />
          <XAxis type="number" xAxisId="pppp" hide />
          <YAxis type="category" dataKey="name" width={110} tick={{ fill: "#8B93A1", fontSize: 12 }} axisLine={{ stroke: LINE }} />
          <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }} labelStyle={{ color: PAPER }} />
          <Legend wrapperStyle={{ fontSize: 11, color: "#8B93A1" }} />
          <RBar xAxisId="freq" dataKey="Frequency %" fill={AMBER} radius={[0, 4, 4, 0]} barSize={14} />
          <Line xAxisId="pppp" dataKey="PPPP" stroke={TEAL} strokeWidth={2} dot={{ r: 3, fill: TEAL }} />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}

function MetricBarList({ title, items, color = AMBER }) {
  if (!items.length) return null;
  return (
    <div data-no-split="true" style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, flex: "1 1 320px" }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, marginBottom: 4 }}>{title}</div>
      <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr 0.55fr 0.55fr", padding: "8px 0 6px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
        <div>Name</div><div>Frequency</div><div>PPPP</div><div>Open</div>
      </div>
      {items.map((it, i) => (
        <div key={i} style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr 0.55fr 0.55fr", padding: "9px 0", alignItems: "center", borderBottom: i < items.length - 1 ? `1px solid ${LINE}` : "none", fontSize: 13 }}>
          <div style={{ color: it.name === "Autres" ? "#8B93A1" : PAPER }}>{it.name}</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 60 }}><Bar pct={it.freq} /></div>
            <span style={{ fontFamily: "ui-monospace, monospace", fontSize: 12, color: "#8B93A1" }}>{it.freq.toFixed(0)}%</span>
          </div>
          <div style={{ fontFamily: "ui-monospace, monospace", color, fontWeight: 700 }}>{it.pppp !== null && it.pppp !== undefined ? it.pppp.toFixed(2) : "–"}</div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: it.open !== null && it.open !== undefined ? TEAL : "#5C6470" }}>{it.open !== null && it.open !== undefined ? it.open.toFixed(0) + "%" : "–"}</div>
        </div>
      ))}
    </div>
  );
}

function SimpleBarChart({ title, data, color = AMBER }) {
  if (!data.length) return null;
  const chartData = data.map(d => ({ name: d.name, "Frequency %": Number((d.freq ?? (100 * d.value)).toFixed(1)) }));
  return (
    <div data-no-split="true" style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "18px 10px 8px 0", flex: "1 1 320px" }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, padding: "0 18px 12px" }}>{title}</div>
      <ResponsiveContainer width="100%" height={Math.max(160, chartData.length * 36)}>
        <ComposedChart data={chartData} layout="vertical" margin={{ top: 0, right: 24, left: 0, bottom: 0 }}>
          <CartesianGrid stroke={LINE} horizontal={false} />
          <XAxis type="number" tick={{ fill: "#5C6470", fontSize: 11 }} tickFormatter={v => v + "%"} axisLine={{ stroke: LINE }} />
          <YAxis type="category" dataKey="name" width={110} tick={{ fill: "#8B93A1", fontSize: 12 }} axisLine={{ stroke: LINE }} />
          <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }} labelStyle={{ color: PAPER }} />
          <RBar dataKey="Frequency %" fill={color} radius={[0, 4, 4, 0]} barSize={16} />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}

function DonutCard({ title, data, note, unit }) {
  return (
    <div data-no-split="true" style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, display: "flex", gap: 18, alignItems: "center", flex: "1 1 280px" }}>
      <DonutChart data={data} unit={unit} />
      <div style={{ flex: 1 }}>
        <div style={{ fontSize: 12, color: "#5C6470", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>{title}</div>
        {data.length ? <DonutLegend data={data} /> : <span style={{ fontSize: 12.5, color: "#5C6470" }}>{note || "Not enough data."}</span>}
      </div>
    </div>
  );
}

function tagBreakdown(source) {
  const map = {};
  source.forEach(p => {
    Object.entries(p.tags).forEach(([tag, v]) => {
      if (!v || EXCLUDED_FROM_BREAKDOWN.has(tag)) return;
      if (!map[tag]) map[tag] = { count: 0, pts: 0, plays: [] };
      map[tag].count += 1;
      map[tag].pts += playPoints(p.tags);
      map[tag].plays.push(p);
    });
  });
  return Object.entries(map)
    .map(([tag, v]) => ({ tag, count: v.count, freq: source.length ? (100 * v.count) / source.length : 0, pppp: v.count ? v.pts / v.count : 0, open: openPct(v.plays) }))
    .sort((a, b) => b.count - a.count);
}

// Bloc complet Attaque / Defense (Plays, Playtypes, Shooting Selection, Erreurs
// défensives) — utilisé pour un joueur comme pour l'équipe entière.
function OffenseDefenseBreakdown({ off, def, detailTables = true, categories }) {
  const cats = categories || currentTagCategories();
  const offPlaysDonut = useMemo(() => categoryBreakdown(off, PLAYS_LIST(cats)).sort((a, b) => b.count - a.count).map(g => ({ name: g.label, value: g.count, freq: g.freq, pppp: g.pppp, open: g.open })), [off]);
  const offPlaysList = useMemo(() => topCategoryBucket(off, PLAYS_LIST(cats), 9), [off]);
  const offPlaytypesList = useMemo(() => topBucket(off, PLAYTYPES_LIST(cats), 8, false), [off]);
  const offShooting = useMemo(() => shootingSelection(off, cats), [off]);

  const defPlaysList = useMemo(() => topCategoryBucket(def, PLAYS_LIST(cats), 9), [def]);
  const defPlaytypesList = useMemo(() => topBucket(def, PLAYTYPES_LIST(cats), 8, false), [def]);
  const defShooting = useMemo(() => shootingSelection(def, cats), [def]);
  const defMistakes = useMemo(() => {
    const b = topBucket(def, DEFENSIVE_MISTAKES_LIST(cats), 9, false);
    // Le fichier original calcule la part de chaque erreur PARMI les erreurs (total = 100%
    // sur ces 9 catégories), pas par rapport à l'ensemble des possessions défendues.
    return b.map((d, i) => ({ name: d.name, value: d.value, color: CHART_COLORS[i % CHART_COLORS.length] }));
  }, [def]);

  const offTagStats = useMemo(() => tagBreakdown(off), [off]);
  const defTagStats = useMemo(() => tagBreakdown(def), [def]);

  // Screen defense subie (attaque : quelle couverture l'adversaire utilise contre nos
  // écrans) et défense d'écran pratiquée (défense : quelle couverture on utilise nous-mêmes).
  const offScreenDef = useMemo(() => categoryBreakdown(off, SCREEN_DEFENSE_LIST(cats)).sort((a, b) => b.count - a.count).map(g => ({ name: g.label, value: g.count, freq: g.freq, pppp: g.pppp, open: g.open })), [off]);
  const defScreenDef = useMemo(() => categoryBreakdown(def, SCREEN_DEFENSE_LIST(cats)).sort((a, b) => b.count - a.count).map(g => ({ name: g.label, value: g.count, freq: g.freq, pppp: g.pppp, open: g.open })), [def]);

  // Spacing joué au moment des écrans, attaque et défense.
  const offSpacing = useMemo(() => categoryBreakdown(off, SPACING_LIST(cats)).sort((a, b) => b.count - a.count).map(g => ({ name: g.label, value: g.count, freq: g.freq, pppp: g.pppp, open: g.open })), [off]);
  const defSpacing = useMemo(() => categoryBreakdown(def, SPACING_LIST(cats)).sort((a, b) => b.count - a.count).map(g => ({ name: g.label, value: g.count, freq: g.freq, pppp: g.pppp, open: g.open })), [def]);

  // Custom categories (créées dans Settings, au-delà des catégories intégrées) —
  // affichées automatiquement ici, attaque et défense, sans rien coder de plus.
  const BUILTIN_CATEGORIES = new Set(["Player", "Playtypes", "Plays", "Shot selection", "Defensive mistakes", "Screen defense", "Spacing", "Shot zone", "Results & misc."]);
  const customCategoryNames = Object.keys(cats).filter(n => !BUILTIN_CATEGORIES.has(n));
  const customOff = useMemo(() => customCategoryNames.map(name => ({ name, items: topBucket(off, categoryTags(name, cats), 8, false) })), [off, customCategoryNames.join(",")]);
  const customDef = useMemo(() => customCategoryNames.map(name => ({ name, items: topBucket(def, categoryTags(name, cats), 8, false) })), [def, customCategoryNames.join(",")]);

  if (!off.length && !def.length) return <EmptyState text="No action coded yet (Import Match tab)." />;

  return (
    <>
      <SectionTitle eyebrow="Source: coding file" title="Offense — how the ball is played" />
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 26 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Plays (game entries)" data={offPlaysDonut.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} />
          <MetricBarList title="Efficiency by play" items={offPlaysList} color={AMBER} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Shooting Selection" data={offShooting} note="Open/Contested tags not present yet." />
          <MetricBarList title="Efficiency by playtype" items={offPlaytypesList} color={AMBER} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Screen defense faced (opponent coverage)" data={offScreenDef.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} note="No screen coverage tag detected." />
          <MetricBarList title="Efficiency by coverage faced" items={offScreenDef} color={AMBER} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Spacing played on screens" data={offSpacing.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} note="No spacing tag detected." />
          <MetricBarList title="Efficiency by spacing" items={offSpacing} color={AMBER} />
        </div>
      </div>

      <SectionTitle eyebrow="Source: coding file" title="Defense — what we defend against" />
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 26 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Shooting Selection (defense)" data={defShooting} note="Open/Contested tags not present yet." />
          <MetricBarList title="Playtypes defended — efficiency" items={defPlaytypesList} color={TEAL} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Defensive mistakes" data={defMistakes} note="No defensive mistake tag detected." />
          <MetricBarList title="Plays defended — efficiency" items={defPlaysList} color={TEAL} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Screen defense played" data={defScreenDef.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} note="No screen coverage tag detected." />
          <MetricBarList title="Efficiency by coverage used" items={defScreenDef} color={TEAL} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 320px" }}>
          <DonutCard title="Spacing faced on screens" data={defSpacing.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} note="No spacing tag detected." />
          <MetricBarList title="Efficiency by spacing faced" items={defSpacing} color={TEAL} />
        </div>
      </div>

      {customCategoryNames.length > 0 && (
        <>
          <SectionTitle eyebrow="Custom categories" title="Settings → other categories" />
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 26 }}>
            {customOff.filter(c => c.items.length).map(c => {
              const style = chartStyleFor(c.name);
              return (
                <React.Fragment key={"off-" + c.name}>
                  {(style === "simple" || style === "both") && <DonutCard title={`${c.name} — offense`} data={c.items.map((it, i) => ({ ...it, color: CHART_COLORS[i % CHART_COLORS.length] }))} />}
                  {(style === "detailed" || style === "both") && <MetricBarList title={`${c.name} — offense`} items={c.items} color={AMBER} />}
                </React.Fragment>
              );
            })}
            {customDef.filter(c => c.items.length).map(c => {
              const style = chartStyleFor(c.name);
              return (
                <React.Fragment key={"def-" + c.name}>
                  {(style === "simple" || style === "both") && <DonutCard title={`${c.name} — defense`} data={c.items.map((it, i) => ({ ...it, color: CHART_COLORS[i % CHART_COLORS.length] }))} />}
                  {(style === "detailed" || style === "both") && <MetricBarList title={`${c.name} — defense`} items={c.items} color={TEAL} />}
                </React.Fragment>
              );
            })}
          </div>
        </>
      )}

      {detailTables && (
        <>
          <SectionTitle eyebrow="Full detail" title="All tags — offense" />
          <TagTable stats={offTagStats} />
          <div style={{ height: 20 }} />
          <SectionTitle eyebrow="Full detail" title="All tags — defense" />
          <TagTable stats={defTagStats} />
        </>
      )}
    </>
  );
}

// ---------------------------------------------------------------------------
// Hook : stats de box score d'un joueur (source de vérité pour les totaux officiels)
// ---------------------------------------------------------------------------

// Fait remonter le temps de jeu en tête des stats affichées, dès qu'il est présent dans
// le fichier importé (colonne "Playing time", "Minutes", "MIN"...).
function prioritizeLabels(labels) {
  const isPlayingTime = l => /temps\s*de\s*jeu|^min(ute)?s?$/i.test(l.trim());
  return [...labels].sort((a, b) => (isPlayingTime(b) ? 1 : 0) - (isPlayingTime(a) ? 1 : 0));
}

function useBoxScore(playerName, filterKeys) {
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [playerName, filterKeys]);

  async function load() {
    setLoading(true);
    const idx = (await storeGet("boxscore_index")) || [];
    const out = [];
    for (const m of idx) {
      if (filterKeys && !filterKeys.has(matchKey(m.date, m.opponent))) continue;
      const data = await storeGet("boxscore:" + m.id);
      if (!data) continue;
      const row = data.rows.find(r => r.player === playerName);
      if (row) {
        // Possessions de l'ÉQUIPE pour CE match, nécessaires pour le %usage du joueur — priorité
        // à la ligne de totaux d'équipe si détectée, sinon somme des lignes joueurs du match.
        const teamStats = data.teamRow ? data.teamRow.stats : sumRowStats(data.rows);
        const teamPossessions = computeTeamPossessionsFromStats(teamStats);
        // Durée totale du match : 5 joueurs sur le terrain en permanence, donc la somme des
        // minutes de TOUS les joueurs (hors ligne de totaux) équivaut à 5 × la durée réelle du
        // match — plus fiable que de supposer 40 minutes fixes (prolongations, format jeune...).
        const minutesCol = findStatCol(Object.keys(row.stats), STAT_PATTERNS.minutes);
        const playerRows = data.rows.filter(r => r !== data.teamRow);
        const totalPlayerMinutes = minutesCol ? sumStat(playerRows, minutesCol) : null;
        const teamMinutes = totalPlayerMinutes ? totalPlayerMinutes / 5 : null;
        out.push({ date: m.date, opponent: m.opponent, stats: { ...row.stats, ...derivedMatchStats(row.stats, teamPossessions, teamMinutes) } });
      }
    }
    setEntries(out.sort((a, b) => a.date.localeCompare(b.date)));
    setLoading(false);
  }

  const statLabels = filterRedundantRawPctColumns(prioritizeLabels(Array.from(new Set(entries.flatMap(e => Object.keys(e.stats))))));
  const weighted = computeWeightedPlayerPercentages(entries);
  // BUG RÉEL CORRIGÉ (même erreur que côté équipe, signalée par l'utilisateur) : les
  // pourcentages d'un joueur sur plusieurs matchs ne peuvent pas être moyennés match par
  // match — un match à 1 tir tenté compterait autant qu'un match à 15. On recalcule ces
  // libellés précis depuis les comptages sommés (weighted), tout le reste reste une moyenne
  // simple (correcte pour des stats de comptage comme les points ou les rebonds).
  const WEIGHTED_LABELS = {
    "% 2pts": "pct2", "% 2pts (calculated)": "pct2", "% 3pts": "pct3", "% 3pts (calculated)": "pct3",
    "% LF": "pctFT", "% LF (calculated)": "pctFT", "eFG% (calculated)": "efg", "Usage%": "usagePct",
  };
  const averages = {};
  statLabels.forEach(l => {
    const vals = entries.map(e => e.stats[l]).filter(v => v !== undefined);
    const naiveAvg = vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
    // BUG RÉEL CORRIGÉ : si le fichier n'a pas de colonnes détaillées (tirs réussis/manqués
    // séparés) pour recalculer ce pourcentage à partir des totaux, le calcul pondéré renvoie
    // null — dans ce cas on retombe sur la moyenne simple plutôt que d'afficher "–" alors
    // qu'une vraie valeur existait pour ce match (constaté : un joueur à 1 seul match et un
    // vrai pourcentage de tir s'affichait vide).
    if (WEIGHTED_LABELS[l]) { averages[l] = weighted[WEIGHTED_LABELS[l]] ?? naiveAvg; return; }
    averages[l] = naiveAvg;
  });

  return { entries, statLabels, averages, loading, reload: load };
}

// Détection souple d'une colonne de stat par motif (ex. FGM/FG Made/2PM+3PM… selon le nom
// exact utilisé dans le fichier de box score, qu'on ne connaît pas à l'avance).
// Convertit un nom de colonne brut ("Bp", "2Pts réussi"...) tel qu'il vient du fichier importé
// en libellé lisible en anglais, quand la colonne est reconnue par un des motifs STAT_PATTERNS
// ou un alias ajouté manuellement dans Settings. Sinon, renvoie le nom brut inchangé (colonnes
// spécifiques comme un nom de play n'ont pas d'équivalent générique).
function friendlyStatLabel(rawLabel) {
  for (const [key, patterns] of Object.entries(STAT_PATTERNS)) {
    if (patterns.some(p => p.test(String(rawLabel).trim()))) return STAT_KEY_FRIENDLY_NAME[key] || rawLabel;
  }
  const aliases = boxColumnAliases();
  for (const [key, list] of Object.entries(aliases)) {
    if ((list || []).some(a => a.trim().toLowerCase() === String(rawLabel).trim().toLowerCase())) return STAT_KEY_FRIENDLY_NAME[key] || rawLabel;
  }
  return rawLabel;
}

function findStatCol(labels, patterns, key) {
  if (key) {
    const aliases = boxColumnAliases()[key] || [];
    if (aliases.length) {
      const aliasMatch = labels.find(l => aliases.some(a => a.trim().toLowerCase() === String(l).trim().toLowerCase()));
      if (aliasMatch !== undefined) return aliasMatch;
    }
  }
  return labels.find(l => patterns.some(p => p.test(l.trim())));
}

// Correspondances de colonnes de box score ajoutées manuellement par le coach (Settings)
// — vérifiées en priorité, avant les motifs automatiques, pour les cas que la détection
// intégrée ne reconnaît pas. Même schéma que les catégories du fichier de coding.
let BOX_COLUMN_ALIASES = null;
function boxColumnAliases() { return BOX_COLUMN_ALIASES || {}; }
async function loadBoxColumnAliases() {
  const stored = await storeGet("box_column_aliases");
  BOX_COLUMN_ALIASES = stored || {};
}
async function saveBoxColumnAliases(aliases) {
  BOX_COLUMN_ALIASES = aliases;
  await storeSet("box_column_aliases", aliases);
}

// Détection FR + EN : les box scores francophones découpent souvent les tirs en colonnes
// séparées "2Pts Mades" / "2Pts Misseds" plutôt qu'une colonne FGA/FGM unique — et
// utilisent des abréviations comme "Bp" (Balles perdues = pertes de balle), "Ro" (Rebonds
// offensifs), "LF" (Lancers Francs). On détecte chaque brique puis on recombine.
const MADE = "(r[ée]ussis?|made)$";
const MISSED = "(manqu[ée]s?|missed)$";
const STAT_PATTERNS = {
  minutes: [/^Temps\s*de\s*jeu$/i, /^Min(ute)?s?$/i, /^MIN$/i, /^Time$/i],
  made2: [new RegExp(`^2\\s*(pts?|points?)?\\s*${MADE}`, "i"), /^2PM$/i, /^2P\+$/i, /^FGM$/i, /^FG\s*Made$/i],
  missed2: [new RegExp(`^2\\s*(pts?|points?)?\\s*${MISSED}`, "i"), /^2P-$/i],
  made3: [new RegExp(`^3\\s*(pts?|points?)?\\s*${MADE}`, "i"), /^3PM$/i, /^3P\+$/i, /^Threes?\s*Made$/i],
  missed3: [new RegExp(`^3\\s*(pts?|points?)?\\s*${MISSED}`, "i"), /^3P-$/i],
  madeFT: [/^LF\s*R[ée]ussis?$/i, new RegExp(`^L\\.?F\\.?\\s*${MADE}`, "i"), new RegExp(`lancers?\\s*francs?\\s*${MADE}`, "i"), /^FTM$/i, /^FT\s*Made$/i, new RegExp(`^FT\\s*${MADE}`, "i")],
  missedFT: [/^LF\s*Manqu[ée]s?$/i, new RegExp(`^L\\.?F\\.?\\s*${MISSED}`, "i"), new RegExp(`lancers?\\s*francs?\\s*${MISSED}`, "i"), /^FT-$/i, new RegExp(`^FT\\s*${MISSED}`, "i")],
  fgm: [/^FGM$/i, /^FG\s*Made$/i, /^Field\s*Goals?\s*Made$/i],
  // Priorité au sous-total "Tot." des groupes fusionnés (tentatives réelles) — sinon on
  // retomberait sur la sous-colonne "Made" (tirs marqués), ce qui a causé la sous-évaluation
  // massive des possessions.
  fga: [/^Tirs\s*Tot\.?\s*Tot\.?$/i, /^FGA$/i, /^FG\s*Att(empt(s|ed)?)?$/i, /^Field\s*Goals?\s*Att/i, /^Tirs?\s*tent[ée]s?$/i, /^Tirs\s*tot\.?$/i, /^Tirs?\s*totale?s?$/i],
  tpm: [/^3PM$/i, /^3P\s*Made$/i, /^Threes?\s*Made$/i],
  fta: [/^LF\s*Tot\.?$/i, /^FT\s*Tot\.?$/i, /^FTA$/i, /^FT\s*Att(empt(s|ed)?)?$/i, /^Free\s*Throws?\s*Att/i, /^LF\s*tent[ée]s?$/i, /^L\.?F\.?$/i, /^LF$/i],
  ftm: [/^FTM$/i, /^FT\s*Made$/i, /^Free\s*Throws?\s*Made$/i],
  oreb: [/^Reb\s*Off\.?$/i, /^Ro$/i, /^R\.?O\.?$/i, /rebonds?\s*off/i, /^OREB$/i, /^ORB$/i, /^Off\.?$/i],
  reb: [/^Reb\s*Tot\.?$/i, /^Reb$/i, /^Rebonds?$/i, /^R\.?D\.?$/i], // rebonds TOTAUX (offensifs+défensifs) — informatif
  ast: [/^Pad$/i, /^P\.?D\.?$/i, /passes?\s*d[ée]cisives?/i, /^AST$/i, /assists?/i],
  tov: [/^Bp$/i, /^B\.?P\.?$/i, /balles?\s*perdues?/i, /pertes?\s*(de\s*)?balles?/i, /^TOV$/i, /^TO$/i, /turn\s*overs?/i],
  pts: [/^Pts?$/i, /^Points?$/i],
  ftPct: [/^LF\s*%$/i, /^FT\s*%$/i, /^FT%$/i, /%\s*LF$/i],
  tpmPct: [/^3\s*pts?\s*%$/i, /^3P\s*%$/i, /^3PT\s*%$/i],
  twoPct: [/^2\s*pts?\s*%$/i, /^2P\s*%$/i, /^2PT\s*%$/i],
};

// Statistiques mises en avant selon le poste — utilisées pour les totaux officiels d'un
// joueur et pour la comparaison par poste.
const POSITION_FEATURED_STATS = {
  "Point Guard": ["pts", "ast", "tov", "ftPct", "tpmPct"],
  "Forward": ["pts", "oreb", "tov", "ftPct", "tpmPct"],
  "Big": ["pts", "reb", "tov", "ftPct", "twoPct"],
};
const STAT_KEY_LABEL_FR = { pts: "Pts", ast: "Assists", tov: "Turnovers", ftPct: "% FT", tpmPct: "% 3PT", oreb: "Off. rebounds", reb: "Rebounds", twoPct: "% 2PT" };

// Noms lisibles pour toutes les clés STAT_PATTERNS — utilisés dans Settings pour laisser
// le coach ajouter/modifier/supprimer les noms de colonnes reconnus pour chaque stat.
const STAT_KEY_FRIENDLY_NAME = {
  minutes: "Minutes / Playing time", made2: "2PT Made", missed2: "2PT Missed", made3: "3PT Made", missed3: "3PT Missed",
  madeFT: "FT Made", missedFT: "FT Missed", fga: "FG Attempted (total)", fta: "FT Attempted (total)",
  tov: "Turnovers", oreb: "Offensive Rebounds", reb: "Rebounds (total)", ast: "Assists", pts: "Points",
  twoPct: "% 2PT (if already in the file)", tpmPct: "% 3PT (if already in the file)", ftPct: "% FT (if already in the file)",
};

// Retire les colonnes de pourcentage BRUTES du fichier (2PTS%, 3PT %…) de la liste des
// statistiques affichées : elles font doublon avec "% 2pts"/"% 3pts"/"% LF" (calculées par
// derivedMatchStats, correctement moyennées sur plusieurs matchs) et créaient une confusion
// réelle — deux colonnes différentes pour la même chose, l'une correcte et l'autre pas
// (constaté : 33% affiché, moyenne naïve fausse de la colonne brute, alors que la version
// normalisée affichait le bon 25%). Une seule version claire désormais.
// BUG RÉEL CORRIGÉ : le motif de détection de "FT%" brut ("%\s*LF$") correspond aussi, par
// accident, au nom de NOTRE PROPRE libellé normalisé "% LF" — le filtre le supprimait donc
// lui-même, faisant disparaître %FT partout (fiche joueur, comparaison par poste). On exempte
// maintenant explicitement nos propres libellés dérivés, jamais concernés par ce filtre.
const NORMALIZED_DERIVED_LABELS = new Set(["% 2pts", "% 2pts (calculated)", "% 3pts", "% 3pts (calculated)", "% LF", "% LF (calculated)"]);
function filterRedundantRawPctColumns(statLabels) {
  const rawPctPatterns = [...STAT_PATTERNS.twoPct, ...STAT_PATTERNS.tpmPct, ...STAT_PATTERNS.ftPct];
  return statLabels.filter(l => NORMALIZED_DERIVED_LABELS.has(l) || !rawPctPatterns.some(p => p.test(String(l).trim())));
}

// Pour un joueur donné, retrouve — parmi SES colonnes réellement présentes dans le fichier
// importé — le nom exact de colonne correspondant à chaque stat mise en avant pour son poste.
function featuredStatsForPosition(position, statLabels) {
  const keys = POSITION_FEATURED_STATS[position] || ["pts", "tov"];
  // Même correctif que PositionComparisonTable : priorise la colonne normalisée ("% 3pts")
  // à la colonne brute du fichier, qui peut varier d'un fichier importé à l'autre.
  const NORMALIZED_PCT_LABEL = { twoPct: "% 2pts", tpmPct: "% 3pts", ftPct: "% LF" };
  return keys.map(key => {
    if (NORMALIZED_PCT_LABEL[key] && statLabels.includes(NORMALIZED_PCT_LABEL[key])) {
      return { key, fallbackLabel: STAT_KEY_LABEL_FR[key] || key, label: NORMALIZED_PCT_LABEL[key] };
    }
    return { key, fallbackLabel: STAT_KEY_LABEL_FR[key] || key, label: findStatCol(statLabels, STAT_PATTERNS[key] || [], key) };
  });
}

// Cherche une valeur numérique dans la ligne d'un joueur pour une catégorie donnée, en
// testant toutes les colonnes candidates (utile pour made2/missed2/etc. qui n'ont pas de
// colonne combinée directe).
function sumStat(rows, colName) {
  if (!colName) return null;
  return rows.reduce((s, r) => s + (Number(r.stats[colName]) || 0), 0);
}
// Additionne toutes les colonnes numériques d'un ensemble de lignes joueurs en un seul objet
// de stats — repli utilisé quand le fichier n'a pas de ligne de totaux d'équipe explicite.
function sumRowStats(rows) {
  const out = {};
  for (const r of rows) {
    for (const [k, v] of Object.entries(r.stats)) {
      const n = Number(v);
      if (!Number.isNaN(n)) out[k] = (out[k] || 0) + n;
    }
  }
  return out;
}

// Estimation standard des possessions d'ÉQUIPE (Dean Oliver) à partir d'un objet de stats
// brutes quelconque (ligne de totaux d'équipe, ou somme des joueurs) : FGA - OREB + TOV +
// 0.44*FTA. Réutilisée à la fois pour les stats avancées d'équipe et pour le %usage individuel.
function computeTeamPossessionsFromStats(statsObj) {
  if (!statsObj) return null;
  const labels = Object.keys(statsObj);
  const get = (patterns) => { const l = findStatCol(labels, patterns); return l !== undefined ? statsObj[l] : undefined; };
  const made2 = get(STAT_PATTERNS.made2), missed2 = get(STAT_PATTERNS.missed2);
  const made3 = get(STAT_PATTERNS.made3), missed3 = get(STAT_PATTERNS.missed3);
  const fga = get(STAT_PATTERNS.fga) ?? ((made2 !== undefined || missed2 !== undefined || made3 !== undefined || missed3 !== undefined)
    ? (made2 || 0) + (missed2 || 0) + (made3 || 0) + (missed3 || 0) : undefined);
  const fta = get(STAT_PATTERNS.fta);
  const tov = get(STAT_PATTERNS.tov);
  const oreb = get(STAT_PATTERNS.oreb);
  if (fga === undefined || fta === undefined || tov === undefined) return null;
  return fga - (oreb || 0) + tov + 0.44 * fta;
}

// À partir des colonnes brutes d'UN match (ex. "2pts", "3pts Missed"…), calcule les
// pourcentages qui ne sont pas fournis tels quels par le fichier — pour qu'ils soient
// sélectionnables dans les courbes d'évolution même si le club ne les exporte pas.
// teamPossessions (optionnel) : possessions totales de l'ÉQUIPE pour ce même match.
// teamMinutes (optionnel) : durée totale du match, en minutes (ex. 40) — pour ramener le
// %usage aux possessions THÉORIQUES pendant le temps de jeu du joueur, pas tout le match :
// possessions théoriques sur le terrain = (minutes du joueur / durée du match) × possessions
// d'équipe. Le %usage est alors : possessions terminées ÷ possessions théoriques sur le terrain.
function derivedMatchStats(statsObj, teamPossessions, teamMinutes) {
  const labels = Object.keys(statsObj);
  const get = (patterns) => { const l = findStatCol(labels, patterns); return l !== undefined ? statsObj[l] : undefined; };
  const made2 = get(STAT_PATTERNS.made2), missed2 = get(STAT_PATTERNS.missed2);
  const made3 = get(STAT_PATTERNS.made3), missed3 = get(STAT_PATTERNS.missed3);
  const madeFT = get(STAT_PATTERNS.madeFT), missedFT = get(STAT_PATTERNS.missedFT);
  // BUG RÉEL CORRIGÉ : la colonne du fichier est souvent une fraction 0-1 (ex. 0.319 pour
  // 31.9%) alors que le calcul de repli produit un nombre 0-100 — sans cette normalisation,
  // les courbes d'évolution afficheraient une échelle complètement fausse pour ce match précis.
  const normalizePct = (v) => (v === undefined || v === null) ? undefined : (Math.abs(v) <= 1 ? v * 100 : v);
  const directPct2 = normalizePct(get(STAT_PATTERNS.twoPct)), directPct3 = normalizePct(get(STAT_PATTERNS.tpmPct)), directPctFT = normalizePct(get(STAT_PATTERNS.ftPct));
  const derived = {};
  if (directPct2 !== undefined) derived["% 2pts"] = directPct2;
  else if (made2 !== undefined && missed2 !== undefined && made2 + missed2 > 0)
    derived["% 2pts (calculated)"] = (100 * made2) / (made2 + missed2);
  if (directPct3 !== undefined) derived["% 3pts"] = directPct3;
  else if (made3 !== undefined && missed3 !== undefined && made3 + missed3 > 0)
    derived["% 3pts (calculated)"] = (100 * made3) / (made3 + missed3);
  if (directPctFT !== undefined) derived["% LF"] = directPctFT;
  else if (madeFT !== undefined && missedFT !== undefined && madeFT + missedFT > 0)
    derived["% LF (calculated)"] = (100 * madeFT) / (madeFT + missedFT);
  // BUG RÉEL CORRIGÉ (même incohérence que dans computeWeightedPlayerPercentages) : le
  // numérateur (tirs réussis) est dérivé de "2pts/3pts Made", donc le dénominateur doit venir
  // de la même décomposition — pas de la colonne "FG Attempted" du fichier, qui peut différer.
  const hasBreakdown2 = made2 !== undefined || missed2 !== undefined || made3 !== undefined || missed3 !== undefined;
  const fgm = (made2 || 0) + (made3 || 0), fga = hasBreakdown2 ? (made2 || 0) + (missed2 || 0) + (made3 || 0) + (missed3 || 0) : get(STAT_PATTERNS.fga);
  if (fga > 0) derived["eFG% (calculated)"] = (100 * (fgm + 0.5 * (made3 || 0))) / fga;

  // Possessions terminées par le joueur : tirs tentés + pertes de balle + 0.44 × lancers francs
  // tentés. Et %usage : cette part rapportée aux possessions THÉORIQUES de l'équipe pendant le
  // temps de jeu du joueur (pas tout le match, sauf s'il a joué la totalité des minutes).
  const fta = get(STAT_PATTERNS.fta) ?? ((madeFT !== undefined || missedFT !== undefined) ? (madeFT || 0) + (missedFT || 0) : undefined);
  const tov = get(STAT_PATTERNS.tov);
  const playerMinutes = get(STAT_PATTERNS.minutes);
  if (fga !== undefined && tov !== undefined && fta !== undefined) {
    const endedPoss = fga + tov + 0.44 * fta;
    derived["Ended possessions"] = endedPoss;
    if (teamPossessions && teamMinutes && playerMinutes !== undefined) {
      const theoreticalPossOnCourt = (playerMinutes / teamMinutes) * teamPossessions;
      if (theoreticalPossOnCourt > 0) {
        derived["Theoretical possessions"] = theoreticalPossOnCourt;
        derived["Usage%"] = (100 * endedPoss) / theoreticalPossOnCourt;
      }
    }
  }
  return derived;
}

// Charge les box scores bruts (toutes les lignes de tous les joueurs, par match) pour
// calculer des statistiques d'ÉQUIPE — possessions, Four Factors, ORTG, DRTG.
// BUG RÉEL CORRIGÉ (signalé par l'utilisateur) : on ne peut pas faire la moyenne de
// pourcentages calculés match par match — un match à 2 tirs tentés compterait autant qu'un
// match à 20 dans la moyenne, ce qui fausse complètement le résultat. La bonne méthode :
// additionner les comptages bruts (réussis, manqués, pertes de balle, possessions...) sur
// TOUS les matchs d'abord, puis calculer UN SEUL pourcentage global à partir de ces totaux.
function computeWeightedTeamPercentages(perMatch) {
  const sum = (key) => perMatch.reduce((s, m) => s + (m[key] !== null && m[key] !== undefined ? m[key] : 0), 0);
  const has = (key) => perMatch.some(m => m[key] !== null && m[key] !== undefined);

  const sumMade2 = sum("made2"), sumMissed2 = sum("missed2");
  const pct2 = (has("made2") && has("missed2") && sumMade2 + sumMissed2 > 0) ? (100 * sumMade2) / (sumMade2 + sumMissed2) : null;

  const sumMade3 = sum("made3"), sumMissed3 = sum("missed3");
  const pct3 = (has("made3") && has("missed3") && sumMade3 + sumMissed3 > 0) ? (100 * sumMade3) / (sumMade3 + sumMissed3) : null;

  const sumMadeFT = sum("madeFT"), sumMissedFT = sum("missedFT");
  const pctFT = (has("madeFT") && has("missedFT") && sumMadeFT + sumMissedFT > 0) ? (100 * sumMadeFT) / (sumMadeFT + sumMissedFT) : null;

  const sumFgm = sum("fgm"), sumTpm = sum("tpm"), sumFga = sum("fga");
  const efg = (has("fgm") && has("tpm") && sumFga > 0) ? (100 * (sumFgm + 0.5 * sumTpm)) / sumFga : null;

  const sumTov = sum("tov"), sumPoss = sum("poss");
  const tovPct = (has("tov") && has("poss") && sumPoss > 0) ? (100 * sumTov) / sumPoss : null;

  const sumOreb = sum("oreb"), sumMissedFG = sumMissed2 + sumMissed3;
  const orebOpportunities = sumMissedFG + 0.44 * sumMissedFT;
  const orebPct = (has("oreb") && (has("missed2") || has("missed3")) && has("missedFT") && orebOpportunities > 0) ? (100 * sumOreb) / orebOpportunities : null;

  const sumAst = sum("ast");
  const astOpportunities = sumFgm + 0.44 * sumMadeFT;
  const astPct = (has("ast") && has("fgm") && has("madeFT") && astOpportunities > 0) ? (100 * sumAst) / astOpportunities : null;

  return { pct2, pct3, pctFT, efg, tovPct, orebPct, astPct };
}

// Même correctif, au niveau INDIVIDUEL cette fois : les pourcentages d'un joueur sur
// plusieurs matchs (%2pts, %3pts, %FT, eFG%, Usage%) ne doivent pas non plus être moyennés
// match par match — on part des colonnes brutes de chaque box score, on additionne les
// comptages, puis on calcule un seul pourcentage global.
function computeWeightedPlayerPercentages(entries) {
  const get = (statsObj, patterns) => {
    const l = findStatCol(Object.keys(statsObj), patterns);
    return l !== undefined ? Number(statsObj[l]) : undefined;
  };
  let sumMade2 = 0, sumMissed2 = 0, sumMade3 = 0, sumMissed3 = 0, sumMadeFT = 0, sumMissedFT = 0, sumFga = 0, sumFgm = 0, sumTpm = 0;
  let hasMade2 = false, hasMissed2 = false, hasMade3 = false, hasMissed3 = false, hasMadeFT = false, hasMissedFT = false, hasFga = false;
  let sumEndedPoss = 0, sumTheoPoss = 0, hasUsage = false;

  for (const e of entries) {
    const s = e.stats;
    const made2 = get(s, STAT_PATTERNS.made2), missed2 = get(s, STAT_PATTERNS.missed2);
    const made3 = get(s, STAT_PATTERNS.made3), missed3 = get(s, STAT_PATTERNS.missed3);
    const madeFT = get(s, STAT_PATTERNS.madeFT), missedFT = get(s, STAT_PATTERNS.missedFT);
    // BUG RÉEL CORRIGÉ (eFG% incohérent avec %2pts/%3pts, constaté sur une vraie fiche joueur) :
    // le numérateur (tirs réussis) est TOUJOURS dérivé de "2pts/3pts Made", mais le
    // dénominateur (tirs tentés) préférait la colonne "FG Attempted" du fichier quand elle
    // existait — si cette colonne ne correspond pas exactement à (2pts tentés + 3pts tentés)
    // dans le fichier d'origine, le eFG% se retrouve calculé avec un numérateur et un
    // dénominateur incohérents entre eux. On priorise maintenant la somme dérivée de la même
    // décomposition que le numérateur, et on ne retombe sur la colonne directe du fichier que
    // si cette décomposition est totalement absente.
    const hasBreakdown = made2 !== undefined || missed2 !== undefined || made3 !== undefined || missed3 !== undefined;
    const fga = hasBreakdown ? (made2 || 0) + (missed2 || 0) + (made3 || 0) + (missed3 || 0) : get(s, STAT_PATTERNS.fga);
    if (made2 !== undefined) { sumMade2 += made2; hasMade2 = true; }
    if (missed2 !== undefined) { sumMissed2 += missed2; hasMissed2 = true; }
    if (made3 !== undefined) { sumMade3 += made3; hasMade3 = true; }
    if (missed3 !== undefined) { sumMissed3 += missed3; hasMissed3 = true; }
    if (madeFT !== undefined) { sumMadeFT += madeFT; hasMadeFT = true; }
    if (missedFT !== undefined) { sumMissedFT += missedFT; hasMissedFT = true; }
    if (fga !== undefined) {
      sumFga += fga; hasFga = true;
      sumFgm += (made2 || 0) + (made3 || 0);
      sumTpm += (made3 || 0);
    }
    if (s["Ended possessions"] !== undefined && s["Theoretical possessions"] !== undefined) {
      sumEndedPoss += s["Ended possessions"]; sumTheoPoss += s["Theoretical possessions"]; hasUsage = true;
    }
  }

  const pct2 = (hasMade2 && hasMissed2 && sumMade2 + sumMissed2 > 0) ? (100 * sumMade2) / (sumMade2 + sumMissed2) : null;
  const pct3 = (hasMade3 && hasMissed3 && sumMade3 + sumMissed3 > 0) ? (100 * sumMade3) / (sumMade3 + sumMissed3) : null;
  const pctFT = (hasMadeFT && hasMissedFT && sumMadeFT + sumMissedFT > 0) ? (100 * sumMadeFT) / (sumMadeFT + sumMissedFT) : null;
  const efg = (hasFga && sumFga > 0) ? (100 * (sumFgm + 0.5 * sumTpm)) / sumFga : null;
  const usagePct = (hasUsage && sumTheoPoss > 0) ? (100 * sumEndedPoss) / sumTheoPoss : null;

  return { pct2, pct3, pctFT, efg, usagePct };
}

function useTeamAdvancedStats(filterKeys) {
  const [data, setData] = useState({ perMatch: [], columns: {}, loading: true });

  useEffect(() => { load(); }, [filterKeys]);

  async function load() {
    const idx = (await storeGet("boxscore_index")) || [];
    const matches = [];
    let allLabels = [];
    for (const m of idx) {
      if (filterKeys && !filterKeys.has(matchKey(m.date, m.opponent))) continue;
      const rec = await storeGet("boxscore:" + m.id);
      if (!rec) continue;
      allLabels = allLabels.concat(rec.rows.flatMap(r => Object.keys(r.stats)));
      matches.push(rec);
    }
    const uniqueLabels = Array.from(new Set(allLabels));
    const columns = {};
    Object.entries(STAT_PATTERNS).forEach(([key, patterns]) => { columns[key] = findStatCol(uniqueLabels, patterns, key); });

    const perMatch = matches.map(m => {
      // Si une ligne de totaux d'équipe a été détectée dans le fichier (ex. "Swiss National
      // Team"), on l'utilise directement — plus fiable qu'une somme des joueurs, qui peut
      // sous-compter si un joueur n'a pas été reconnu à l'import.
      const s = (col) => {
        if (!col) return null;
        if (m.teamRow) return (m.teamRow.stats[col] !== undefined) ? Number(m.teamRow.stats[col]) : null;
        return sumStat(m.rows, col);
      };
      const made2 = s(columns.made2), missed2 = s(columns.missed2);
      const made3 = s(columns.made3), missed3 = s(columns.missed3);
      const madeFT = s(columns.madeFT), missedFT = s(columns.missedFT);

      // Colonnes combinées si elles existent, sinon reconstruites depuis le détail 2pts/3pts/LF.
      const fgm = s(columns.fgm) ?? ((made2 !== null || made3 !== null) ? (made2 || 0) + (made3 || 0) : null);
      const fga = s(columns.fga) ?? ((made2 !== null || missed2 !== null || made3 !== null || missed3 !== null)
        ? (made2 || 0) + (missed2 || 0) + (made3 || 0) + (missed3 || 0) : null);
      const tpm = s(columns.tpm) ?? made3;
      const ftm = s(columns.ftm) ?? madeFT;
      const fta = s(columns.fta) ?? ((madeFT !== null || missedFT !== null) ? (madeFT || 0) + (missedFT || 0) : null);
      const oreb = s(columns.oreb);
      const reb = s(columns.reb); // rebonds totaux (non séparés off/def) — informatif
      const tov = s(columns.tov);
      const ast = s(columns.ast);
      const pts = s(columns.pts) ?? ((made2 !== null || made3 !== null || madeFT !== null)
        ? 2 * (made2 || 0) + 3 * (made3 || 0) + (madeFT || 0) : null);
      const dreb = (reb !== null && oreb !== null) ? reb - oreb : null; // rebonds défensifs déduits si on a le total ET l'offensif
      // Priorité à la colonne de pourcentage déjà présente dans le fichier (ex. "2PTS%") — mais
      // uniquement si on lit une vraie ligne de totaux d'équipe : additionner des pourcentages
      // joueur par joueur n'aurait aucun sens (ça peut dépasser 100%). Sans ligne de totaux, on
      // recalcule toujours depuis les comptages réussis/manqués, qui eux s'additionnent correctement.
      // BUG RÉEL CORRIGÉ : la colonne du fichier est souvent une fraction 0-1 (ex. 0.319 pour
      // 31.9%) alors que le calcul de repli produit un nombre 0-100 — sans cette normalisation,
      // "% 2pts" s'affichait comme "0.3%" au lieu de "31.9%" dès qu'un fichier avait une ligne
      // de totaux d'équipe avec pourcentages déjà calculés (constaté sur un vrai fichier importé).
      const normalizePct = (v) => (v === null || v === undefined) ? null : (Math.abs(v) <= 1 ? v * 100 : v);
      const directPct2 = m.teamRow ? normalizePct(s(columns.twoPct)) : null;
      const directPct3 = m.teamRow ? normalizePct(s(columns.tpmPct)) : null;
      const directPctFT = m.teamRow ? normalizePct(s(columns.ftPct)) : null;
      const pct2 = directPct2 ?? ((made2 !== null && missed2 !== null && made2 + missed2 > 0) ? (100 * made2) / (made2 + missed2) : null);
      const pct3 = directPct3 ?? ((made3 !== null && missed3 !== null && made3 + missed3 > 0) ? (100 * made3) / (made3 + missed3) : null);
      const pctFT = directPctFT ?? ((madeFT !== null && missedFT !== null && madeFT + missedFT > 0) ? (100 * madeFT) / (madeFT + missedFT) : null);

      // Estimation standard des possessions (Dean Oliver) : FGA - OREB + TOV + 0.44*FTA.
      // Si seul le total des rebonds est disponible (pas de split offensif/défensif), on
      // calcule une version approchée sans le terme OREB (approxPoss = true dans ce cas).
      const approxPoss = oreb === null;
      const poss = (fga !== null && tov !== null && fta !== null)
        ? fga - (oreb ?? 0) + tov + 0.44 * fta : null;
      const efg = (fgm !== null && tpm !== null && fga) ? (fgm + 0.5 * tpm) / fga : null;
      const tovPct = (tov !== null && poss) ? tov / poss : null;
      const ftRate = (fta !== null && fga) ? fta / fga : null;
      const ortg = (pts !== null && poss) ? (100 * pts) / poss : null;
      const drtg = (m.opponentScore !== null && m.opponentScore !== undefined && poss) ? (100 * m.opponentScore) / poss : null;
      // % Rebonds offensifs : rebonds offensifs ÷ (tirs manqués + 0.44 × lancers francs
      // manqués) — n'a besoin que des propres stats de l'équipe, pas de celles de l'adversaire.
      const missedFG = (missed2 !== null || missed3 !== null) ? (missed2 || 0) + (missed3 || 0) : null;
      const orebOpportunities = (missedFG !== null && missedFT !== null) ? missedFG + 0.44 * missedFT : null;
      const orebPct = (oreb !== null && orebOpportunities) ? (100 * oreb) / orebOpportunities : null;
      // % Passes décisives : passes décisives ÷ (tirs marqués + 0.44 × lancers francs marqués).
      const astOpportunities = (fgm !== null && madeFT !== null) ? fgm + 0.44 * madeFT : null;
      const astPct = (ast !== null && astOpportunities) ? (100 * ast) / astOpportunities : null;
      return {
        date: m.date, opponent: m.opponent, opponentScore: m.opponentScore, pts, poss, approxPoss, efg, tovPct, oreb, dreb, reb, ftRate, ortg, drtg,
        fga, fta, tov, made2, missed2, made3, missed3, madeFT, missedFT, ast, pct2, pct3, pctFT, orebPct, astPct, fgm, tpm,
      };
    });

    // On ne signale une catégorie comme "manquante" que si NI la colonne combinée NI le
    // détail correspondant n'ont été trouvés — sinon la recombinaison a déjà fait le travail.
    const derivedOk = {
      fgm: !!columns.fgm || !!columns.made2 || !!columns.made3,
      fga: !!columns.fga || !!columns.made2 || !!columns.missed2 || !!columns.made3 || !!columns.missed3,
      tpm: !!columns.tpm || !!columns.made3,
      fta: !!columns.fta || !!columns.madeFT || !!columns.missedFT,
      ftm: !!columns.ftm || !!columns.madeFT,
      oreb: !!columns.oreb,
      tov: !!columns.tov,
      pts: !!columns.pts || !!columns.made2 || !!columns.made3 || !!columns.madeFT,
    };
    const missing = Object.entries(derivedOk).filter(([, ok]) => !ok).map(([k]) => k);

    setData({ perMatch, columns, missing, rawLabels: uniqueLabels, loading: false });
  }

  return data;
}

function useAllBoxScores(filterKeys) {
  const [byPlayer, setByPlayer] = useState({});
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [filterKeys]);

  async function load() {
    setLoading(true);
    const idx = (await storeGet("boxscore_index")) || [];
    const map = {};
    for (const m of idx) {
      if (filterKeys && !filterKeys.has(matchKey(m.date, m.opponent))) continue;
      const data = await storeGet("boxscore:" + m.id);
      if (!data) continue;
      data.rows.forEach(row => {
        if (!map[row.player]) map[row.player] = [];
        map[row.player].push({ date: m.date, opponent: m.opponent, season: m.season || null, stats: { ...row.stats, ...derivedMatchStats(row.stats) } });
      });
    }
    const result = {};
    Object.entries(map).forEach(([player, entries]) => {
      const statLabels = filterRedundantRawPctColumns(prioritizeLabels(Array.from(new Set(entries.flatMap(e => Object.keys(e.stats))))));
      const weighted = computeWeightedPlayerPercentages(entries);
      const WEIGHTED_LABELS = {
        "% 2pts": "pct2", "% 2pts (calculated)": "pct2", "% 3pts": "pct3", "% 3pts (calculated)": "pct3",
        "% LF": "pctFT", "% LF (calculated)": "pctFT", "eFG% (calculated)": "efg", "Usage%": "usagePct",
      };
      const averages = {};
      statLabels.forEach(l => {
        const vals = entries.map(e => e.stats[l]).filter(v => v !== undefined);
        const naiveAvg = vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
        if (WEIGHTED_LABELS[l]) { averages[l] = weighted[WEIGHTED_LABELS[l]] ?? naiveAvg; return; }
        averages[l] = naiveAvg;
      });
      // Détection heuristique de la colonne "points" pour le classement d'équipe.
      const ptsLabel = statLabels.find(l => /^pts?$|^points?$/i.test(l.trim()));
      result[player] = { games: entries.length, statLabels, averages, ptsLabel, entries };
    });
    setByPlayer(result);
    setLoading(false);
  }

  return { byPlayer, loading, reload: load };
}

function PlayerAvatar({ playerName, size = 34, editable = false }) {
  const [photo, setPhoto] = useState(undefined); // undefined = loading, null = none
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const inputRef = useRef();

  useEffect(() => { load(); }, [playerName]);
  async function load() {
    try {
      const r = await storeGet("photo:" + playerName);
      setPhoto(r || null);
    } catch { setPhoto(null); }
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setErr(""); setBusy(true);
    try {
      const dataUrl = await fileToResizedDataURL(file);
      await storeSet("photo:" + playerName, dataUrl);
      setPhoto(dataUrl);
    } catch (ex) { setErr("Import de la photo impossible."); }
    setBusy(false);
    if (inputRef.current) inputRef.current.value = "";
  }

  const initials = playerName.slice(0, 2).toUpperCase();

  return (
    <div style={{ position: "relative", width: size, height: size, flexShrink: 0 }}>
      <div style={{
        width: size, height: size, borderRadius: size > 60 ? 16 : 8, overflow: "hidden",
        background: PANEL2, border: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center",
      }}>
        {photo ? (
          <img src={photo} alt={playerName} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
        ) : (
          <span style={{ fontFamily: "ui-monospace, monospace", fontSize: size * 0.34, color: AMBER, fontWeight: 700 }}>{initials}</span>
        )}
      </div>
      {editable && (
        <>
          <button
            onClick={(e) => { e.stopPropagation(); inputRef.current && inputRef.current.click(); }}
            title="Change photo"
            style={{
              position: "absolute", bottom: -4, right: -4, width: 22, height: 22, borderRadius: "50%",
              background: AMBER, border: `2px solid ${INK}`, color: "#1A1300", display: "flex", alignItems: "center", justifyContent: "center",
              cursor: "pointer", padding: 0,
            }}
          >
            <Plus size={12} strokeWidth={3} />
          </button>
          <input ref={inputRef} type="file" accept="image/*" onChange={handleFile} style={{ display: "none" }} />
        </>
      )}
      {busy && <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.5)", borderRadius: size > 60 ? 16 : 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, color: PAPER }}>…</div>}
      {err && <div style={{ position: "absolute", top: "100%", left: 0, marginTop: 4, fontSize: 10, color: RED, whiteSpace: "nowrap" }}>{err}</div>}
    </div>
  );
}

function StatPill({ label, value, sub, tone = "amber" }) {
  const color = tone === "amber" ? AMBER : tone === "teal" ? TEAL : tone === "red" ? RED : PAPER;
  return (
    <div style={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, padding: "14px 16px", minWidth: 120 }}>
      <div className="keep-color" style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 26, fontWeight: 700, color }}>{value}</div>
      <div style={{ fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase", color: "#8B93A1", marginTop: 4 }}>{label}</div>
      {sub && <div style={{ fontSize: 11, color: "#5C6470", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function Bar({ pct, tone = "amber" }) {
  const color = tone === "amber" ? AMBER : tone === "teal" ? TEAL : RED;
  return (
    <div className="keep-color" style={{ background: "#0C0F14", borderRadius: 999, height: 6, width: "100%", overflow: "hidden" }}>
      <div className="keep-color" style={{ width: `${Math.max(0, Math.min(100, pct))}%`, background: color, height: "100%" }} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [loading, setLoading] = useState(true);
  const [team, setTeam] = useState(null); // {id, name, logo, ...} — équipe active
  const [session, setSession] = useState(null); // {name, role}
  const [roster, setRoster] = useState([]);
  const [matchesIndex, setMatchesIndex] = useState([]); // [{id,date,opponent,playsCount}]
  const [matches, setMatches] = useState({}); // id -> {plays}
  const [boxScoreIndex, setBoxScoreIndex] = useState([]); // [{id,date,opponent,matchedCount}]
  const [tab, setTab] = useState("home");
  const [homeNav, setHomeNav] = useState(null); // { playerSubtab } ou { scoutingSubtab, scoutingTeam } — consommé une fois puis remis à null
  const [selectedPlayer, setSelectedPlayer] = useState(null);
  const [selectedMatchKeys, setSelectedMatchKeys] = useState(null); // null = tous les matchs
  const [currentSeason, setCurrentSeason] = useState(defaultSeasonLabel());
  const { config: visibility } = useVisibilityConfig();
  const [seasonFilter, setSeasonFilter] = useState("all"); // "all" ou une saison précise
  const [initializing, setInitializing] = useState(true);

  useEffect(() => {
    (async () => {
      const savedTeamId = await loadLocalActiveTeam();
      if (savedTeamId) {
        const teams = await loadTeams();
        const t = teams.find(x => x.id === savedTeamId);
        if (t) selectTeam(t);
      }
      setInitializing(false);
    })();
  }, []);

  useEffect(() => { if (team) bootstrap(); }, [team]);

  function selectTeam(t) {
    setActiveTeam(t.id); // toutes les clés de stockage sont désormais préfixées pour cette équipe
    setTeam(t);
    saveLocalActiveTeam(t.id);
  }

  async function addPlayer(name, position) {
    const first = name.trim().split(" ")[0];
    const newRoster = [...roster, { id: uid(), name: name.trim(), first, position: position || "" }];
    await storeSet("roster", newRoster);
    setRoster(newRoster);
  }

  async function removePlayer(id, label) {
    await requestDeletion(team.id, team.name, "player", label, { id });
  }

  async function editPlayer(id, updates) {
    const current = roster.find(p => p.id === id);
    if (current && updates.name && updates.name !== current.name) {
      await migratePlayerRename(current.name, updates.name);
    }
    const newRoster = roster.map(p => p.id === id ? { ...p, ...updates } : p);
    await storeSet("roster", newRoster);
    setRoster(newRoster);
  }

  async function bootstrap() {
    await loadTagCategories();
    await loadObservationTagCategories();
    await loadCategoryChartStyles();
    await loadBoxColumnAliases();
    const savedSeason = await storeGet("current_season");
    if (savedSeason) setCurrentSeason(savedSeason); else await storeSet("current_season", currentSeason);
    const r = await storeGet("roster");
    let effectiveRoster = [];
    if (r) {
      const cleaned = r.filter(p => !/^\d+$/.test(String(p.first).trim()));
      setRoster(cleaned);
      effectiveRoster = cleaned;
      if (cleaned.length !== r.length) await storeSet("roster", cleaned); // purge un ancien "joueur" numérique (ligne d'équipe)
    } else {
      // Seule l'équipe nationale suisse démarre avec le roster pré-rempli — les autres
      // équipes (ex. Aurore Vitré) démarrent vides, à construire via "Add a player".
      const seed = team.id === "u16-sui" ? DEFAULT_ROSTER : [];
      setRoster(seed);
      effectiveRoster = seed;
      await storeSet("roster", seed);
    }
    const savedSession = await loadLocalSession();
    if (savedSession) {
      // BUG RÉEL CORRIGÉ : si un joueur est renommé (icône crayon) alors qu'il reste connecté
      // sur son appareil sans se reconnecter, sa session gardait l'ANCIEN nom indéfiniment —
      // toutes ses réponses (Wellness, Training…) continuaient à s'enregistrer sous l'ancien
      // nom, invisibles pour le coach qui consulte désormais le nouveau nom. On resynchronise
      // maintenant TOUJOURS le nom depuis l'id stable du joueur (si disponible), à chaque
      // chargement de l'app — pas seulement pour l'ancienne migration prénom → nom complet.
      if (savedSession.role === "player") {
        const byId = savedSession.id ? effectiveRoster.find(p => p.id === savedSession.id) : null;
        if (byId && byId.name !== savedSession.name) {
          const fixed = { ...savedSession, name: byId.name };
          await saveLocalSession(fixed);
          setSession(fixed);
        } else if (byId) {
          setSession(savedSession);
        } else {
          // Pas d'id en session (ancienne connexion antérieure à ce correctif) — on retombe sur
          // l'ancienne logique de migration prénom seul → nom complet, par nom cette fois.
          const exactMatch = effectiveRoster.find(p => p.name === savedSession.name);
          if (!exactMatch) {
            const byFirstNameOnly = effectiveRoster.find(p => p.first === savedSession.name);
            if (byFirstNameOnly) {
              const fixed = { ...savedSession, name: byFirstNameOnly.name, id: byFirstNameOnly.id };
              await saveLocalSession(fixed);
              setSession(fixed);
            } else {
              setSession(savedSession);
            }
          } else {
            // Rattache l'id maintenant qu'on l'a trouvé par nom, pour que les futurs
            // renommages soient correctement suivis à partir de maintenant.
            const fixed = { ...savedSession, id: exactMatch.id };
            await saveLocalSession(fixed);
            setSession(fixed);
          }
        }
      } else {
        setSession(savedSession);
      }
    }
    const idx = (await storeGet("match_index")) || [];
    setMatchesIndex(idx);
    const loaded = {};
    for (const m of idx) {
      const data = await storeGet("match:" + m.id);
      if (data) loaded[m.id] = data;
    }
    setMatches(loaded);
    const bsIdx = (await storeGet("boxscore_index")) || [];
    setBoxScoreIndex(bsIdx);
    setLoading(false);
  }

  // Combine le filtre de saison (Settings) et la sélection manuelle de matchs (menu déroulant)
  // en un seul ensemble de clés — les deux hooks de box score et allPlays s'appuient dessus.
  const effectiveMatchFilter = useMemo(() => {
    const seasonKeys = seasonFilter === "all" ? null
      : new Set([...matchesIndex, ...boxScoreIndex].filter(m => m.season === seasonFilter).map(m => matchKey(m.date, m.opponent)));
    if (!selectedMatchKeys && !seasonKeys) return null;
    if (!selectedMatchKeys) return seasonKeys;
    if (!seasonKeys) return selectedMatchKeys;
    return new Set([...selectedMatchKeys].filter(k => seasonKeys.has(k)));
  }, [seasonFilter, selectedMatchKeys, matchesIndex, boxScoreIndex]);

  const allPlays = useMemo(() => {
    const out = [];
    for (const m of matchesIndex) {
      const d = matches[m.id];
      if (d) d.plays.forEach(p => out.push({ ...p, matchId: m.id, date: m.date, opponent: m.opponent }));
    }
    if (!effectiveMatchFilter) return out;
    return out.filter(p => effectiveMatchFilter.has(matchKey(p.date, p.opponent)));
  }, [matches, matchesIndex, effectiveMatchFilter]);

  // Liste unifiée de tous les matchs connus (fichier de coding et/ou box score), dédupliquée
  // par date+adversaire, pour le sélecteur de matchs.
  const allMatchOptions = useMemo(() => {
    const map = new Map();
    matchesIndex.forEach(m => {
      const k = matchKey(m.date, m.opponent);
      const existing = map.get(k) || { date: m.date, opponent: m.opponent, hasCoding: false, hasBox: false };
      existing.hasCoding = true;
      map.set(k, existing);
    });
    boxScoreIndex.forEach(m => {
      const k = matchKey(m.date, m.opponent);
      const existing = map.get(k) || { date: m.date, opponent: m.opponent, hasCoding: false, hasBox: false };
      existing.hasBox = true;
      map.set(k, existing);
    });
    return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
  }, [matchesIndex, boxScoreIndex]);

  if (initializing) {
    return (
      <div style={{ minHeight: "100vh", background: INK, color: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "ui-monospace, monospace" }}>
        loading…
      </div>
    );
  }

  if (!team) {
    return <TeamSelectScreen onSelectTeam={selectTeam} />;
  }

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", background: INK, color: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "ui-monospace, monospace" }}>
        loading…
      </div>
    );
  }

  if (!session) {
    return <LoginScreen roster={roster} onLogin={async (s) => { await saveLocalSession(s); setSession(s); }} />;
  }

  return (
    <div style={{ minHeight: "100vh", width: "100%", overflowX: "hidden", background: INK, color: PAPER, fontFamily: "'Inter', -apple-system, Segoe UI, sans-serif" }}>
      <style>{`
        html, body { overflow-x: hidden; max-width: 100%; background: ${INK}; margin: 0; padding: 0; height: 100%; }
        #root { background: ${INK}; min-height: 100vh; min-height: 100dvh; }
        .print-only {
          position: absolute; left: -9999px; top: 0; width: 900px; height: auto; overflow: visible;
        }
        @media print {
          body * { visibility: hidden; }
          .print-only, .print-only * { visibility: visible !important; }
          .print-only { position: absolute; left: 0; top: 0; width: 100%; }
          .printable-root, .printable-root * {
            background: #fff !important; color: #111 !important; border-color: #ddd !important;
          }
        }
      `}</style>
      <TopBar
        session={session}
        team={team}
        onSwitchAccount={async () => { await clearLocalSession(); setSession(null); }}
        onSwitchTeam={async () => {
          await clearLocalActiveTeam();
          await clearLocalSession();
          setActiveTeam(null);
          setTeam(null); setSession(null); setRoster(DEFAULT_ROSTER);
          setMatchesIndex([]); setMatches({}); setBoxScoreIndex([]);
          setSelectedPlayer(null); setSelectedMatchKeys(null); setLoading(true);
        }}
        isCoach={session.role === "coach"}
      />
      <div className="screen-only" style={{ maxWidth: 1180, margin: "0 auto", padding: "28px 20px 96px" }}>
        {tab === "home" && (
          <HomeTab
            session={session}
            isCoach={session.role === "coach"}
            playerName={session.role === "player" ? session.name : null}
            allPlays={allPlays}
            roster={roster}
            matchFilter={effectiveMatchFilter}
            team={team}
            goToPlayer={(subtab) => { setSelectedPlayer(session.name); setHomeNav({ playerSubtab: subtab }); setTab("players"); }}
            goToScouting={(subtab, teamName) => { setHomeNav({ scoutingSubtab: subtab, scoutingTeam: teamName }); setTab("scouting"); }}
            goToTeam={() => setTab("team")}
            goToPlanning={() => setTab("planning")}
            visibility={visibility}
          />
        )}
        {tab === "import" && session.role === "coach" && (
          <ImportTab
            roster={roster}
            onImported={async (parsed, meta) => {
              const id = uid();
              const record = { id, date: meta.date, opponent: meta.opponent, season: currentSeason, plays: parsed.plays };
              await storeSet("match:" + id, record);
              if (parsed.fileDataUrl) await storeSet("match_file:" + id, { name: parsed.fileName || "match.xlsx", dataUrl: parsed.fileDataUrl });
              const newIdx = [...matchesIndex, { id, date: meta.date, opponent: meta.opponent, season: currentSeason, playsCount: parsed.plays.length }]
                .sort((a, b) => a.date.localeCompare(b.date));
              await storeSet("match_index", newIdx);
              setMatchesIndex(newIdx);
              setMatches(m => ({ ...m, [id]: record }));
              // Le roster se synchronise avec les joueurs détectés dans le fichier importé —
              // aucune liste de noms codée en dur n'est nécessaire.
              const known = new Set(roster.map(p => p.first.toLowerCase()));
              const newPlayers = (parsed.detectedPlayers || []).filter(n => !known.has(String(n).trim().toLowerCase()));
              if (newPlayers.length) {
                const additions = newPlayers.map(n => ({ id: uid(), name: n.trim(), first: n.trim(), position: positionOf(n.trim()) }));
                const newRoster = [...roster, ...additions];
                await storeSet("roster", newRoster);
                setRoster(newRoster);
                // Also add them to the "Player" category so they are recognized with
                // certitude au prochain import, pas seulement par la détection par défaut.
                const cats = currentTagCategories();
                const playerList = new Set((cats["Player"] || []).map(t => t.toLowerCase()));
                const toAdd = newPlayers.filter(n => !playerList.has(String(n).trim().toLowerCase()));
                if (toAdd.length) await saveTagCategories({ ...cats, "Player": [...(cats["Player"] || []), ...toAdd.map(n => n.trim())] });
              }
            }}
            matchesIndex={matchesIndex}
            onDeleteMatch={async (id, label) => {
              await requestDeletion(team.id, team.name, "match", label, { id });
            }}
          />
        )}
        {tab === "boxscore" && session.role === "coach" && (
          <BoxScoreTab
            roster={roster}
            index={boxScoreIndex}
            onImported={async (parsed, meta) => {
              const id = uid();
              const record = { id, date: meta.date, opponent: meta.opponent, season: currentSeason, opponentScore: meta.opponentScore ?? null, teamRow: parsed.teamRow || null, rows: parsed.rows };
              await storeSet("boxscore:" + id, record);
              if (parsed.fileDataUrl) await storeSet("boxscore_file:" + id, { name: parsed.fileName || "boxscore.xlsx", dataUrl: parsed.fileDataUrl });
              const newIdx = [...boxScoreIndex, { id, date: meta.date, opponent: meta.opponent, season: currentSeason, matchedCount: parsed.matchedCount }]
                .sort((a, b) => a.date.localeCompare(b.date));
              await storeSet("boxscore_index", newIdx);
              setBoxScoreIndex(newIdx);
            }}
            onDelete={async (id, label) => {
              await requestDeletion(team.id, team.name, "boxscore", label, { id });
            }}
          />
        )}
        {(tab === "players" || tab === "team" || tab === "scouting") && (
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-start" }}>
            <MatchSelector options={allMatchOptions} selectedKeys={selectedMatchKeys} onChange={setSelectedMatchKeys} />
            {(() => {
              const seasons = Array.from(new Set([...matchesIndex, ...boxScoreIndex].map(m => m.season).filter(Boolean)));
              if (!seasons.length) return null;
              return (
                <select value={seasonFilter} onChange={e => setSeasonFilter(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160, marginBottom: 20 }}>
                  <option value="all">All seasons</option>
                  {seasons.sort().reverse().map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              );
            })()}
          </div>
        )}
        {tab === "players" && !selectedPlayer && (session.role === "coach" || visibility.tabs.players) && (
          <PlayersList roster={roster} allPlays={allPlays} onSelect={setSelectedPlayer} matchFilter={effectiveMatchFilter}
            isCoach={session.role === "coach"} onlyOwn={session.role === "player" ? session.name : null}
            onAddPlayer={addPlayer} onRemovePlayer={removePlayer} onEditPlayer={editPlayer} />
        )}
        {tab === "players" && selectedPlayer && (session.role === "coach" || visibility.tabs.players) && (
          <PlayerDetail
            playerName={selectedPlayer}
            allPlays={allPlays}
            roster={roster}
            matchFilter={effectiveMatchFilter}
            onBack={() => setSelectedPlayer(null)}
            isCoach={session.role === "coach"}
            visibility={visibility}
            teamId={team.id}
            teamName={team.name}
            initialSubtab={homeNav?.playerSubtab}
            onEditPlayer={editPlayer}
          />
        )}
        {tab === "players" && session.role === "player" && !visibility.tabs.players && (
          <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5 }}>
            You don't have the permission to see this.
          </div>
        )}
        {tab === "team" && (session.role === "coach" || visibility.tabs.team) && <TeamTab roster={roster} allPlays={allPlays} matchesIndex={matchesIndex} matchFilter={effectiveMatchFilter} isCoach={session.role === "coach"} team={team} visibility={visibility} />}
        {tab === "team" && session.role === "player" && !visibility.tabs.team && (
          <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5 }}>
            You don't have the permission to see this.
          </div>
        )}
        {tab === "scouting" && (session.role === "coach" || visibility.tabs.scouting) && <ScoutingTab isCoach={session.role === "coach"} matchFilter={effectiveMatchFilter} initialSubtab={homeNav?.scoutingSubtab} initialReportTeam={homeNav?.scoutingTeam} />}
        {tab === "scouting" && session.role === "player" && !visibility.tabs.scouting && (
          <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5 }}>
            You don't have the permission to see this.
          </div>
        )}
        {tab === "planning" && (session.role === "coach" || visibility.tabs.planning) && <PlanningTab isCoach={session.role === "coach"} team={team} roster={roster} playerName={session.role === "player" ? session.name : null} />}
        {tab === "planning" && session.role === "player" && !visibility.tabs.planning && (
          <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5 }}>
            You don't have the permission to see this.
          </div>
        )}
        {tab === "backup" && session.role === "coach" && <BackupTab team={team} roster={roster} />}
        {tab === "settings" && session.role === "coach" && (
          <div>
            <SectionTitle eyebrow="Settings" title="Current season" />
            <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 26, display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div>
                <label style={labelStyle}>Season applied to future imports</label>
                <input value={currentSeason} onChange={e => setCurrentSeason(e.target.value)} onBlur={() => storeSet("current_season", currentSeason)} placeholder="ex. 2025-2026" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160 }} />
              </div>
              <div style={{ fontSize: 12, color: "#8B93A1", maxWidth: 420 }}>
                Every match or box score imported is tagged with this season. Change it at the start of a new season — already
                imported matches keep the season they had at import time. You can then filter by
                season in the Players, Team and Scouting tabs.
              </div>
            </div>
            <TagCategoriesSettings roster={roster} title="Column categories — Import Match (coding file)" getCurrent={currentTagCategories} onSave={saveTagCategories} onResetDefault={() => JSON.parse(JSON.stringify(DEFAULT_TAG_CATEGORIES))} />
            <div style={{ height: 34 }} />
            <TagCategoriesSettings roster={roster} title="Column categories — Scouting Observation" getCurrent={currentObservationTagCategories} onSave={saveObservationTagCategories} onResetDefault={() => JSON.parse(JSON.stringify(DEFAULT_TAG_CATEGORIES))} />
            <div style={{ height: 26 }} />
            <BoxColumnAliasesSettings />
          </div>
        )}
      </div>
      <BottomTabBar tab={tab} setTab={(t) => { setHomeNav(null); setTab(t); }} isCoach={session.role === "coach"} visibility={visibility} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Connexion
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Sélection d'équipe (avant la connexion) — chaque équipe a ses propres données,
// totalement séparées, protégées par un code d'accès.
// ---------------------------------------------------------------------------

function TeamSelectScreen({ onSelectTeam }) {
  const [teams, setTeams] = useState([]);
  const [picked, setPicked] = useState(null); // team object en cours de déverrouillage
  const [code, setCode] = useState("");
  const [confirmCode, setConfirmCode] = useState("");
  const [mode, setMode] = useState(null); // "enter" | "create"
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);

  useEffect(() => { loadTeams().then(setTeams); }, []);

  async function pickTeam(team) {
    setPicked(team); setError(""); setCode(""); setConfirmCode("");
    setBusy(true);
    // Le code est lu AVANT toute sélection d'équipe active (TEAM_PREFIX encore vide),
    // donc sur une clé globale non liée à une équipe en particulier.
    const existing = await storeGet("team_access:" + team.id);
    setMode(existing ? "enter" : "create");
    setBusy(false);
  }

  async function submitEnter() {
    if (!code) { setError("Enter the access code."); return; }
    const existing = await storeGet("team_access:" + picked.id);
    if (simpleHash(code) !== existing) { setError("Code incorrect."); return; }
    onSelectTeam(picked);
  }

  async function submitCreate() {
    if (code.length < 4) { setError("Choose a code of at least 4 characters."); return; }
    if (code !== confirmCode) { setError("Les deux codes ne correspondent pas."); return; }
    await storeSet("team_access:" + picked.id, simpleHash(code));
    onSelectTeam(picked);
  }

  if (showAdmin) return <AdminGate onClose={() => setShowAdmin(false)} teams={teams} onTeamsChange={setTeams} />;

  return (
    <div style={{ minHeight: "100vh", background: INK, color: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 20 }}>
      <div style={{ width: 420, maxWidth: "100%" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center", marginBottom: 18 }}>
          <img src="/icon-512.png" alt="" onError={e => { e.target.style.display = "none"; }} style={{ width: 64, height: 64, borderRadius: 14, marginBottom: 10 }} />
          <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontSize: 13, fontWeight: 700, letterSpacing: "0.14em" }}>HOOPTRACK BASKETBALL</div>
        </div>
        <h1 style={{ fontSize: 24, fontWeight: 800, margin: "0 0 22px", letterSpacing: "-0.01em", textAlign: "center" }}>
          {!picked ? "Which team?" : `Code — ${picked.name}`}
        </h1>

        {!picked && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {teams.map(t => (
              <button key={t.id} onClick={() => pickTeam(t)} style={{
                display: "flex", alignItems: "center", gap: 14, width: "100%", padding: "14px 16px",
                background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, color: PAPER, cursor: "pointer", textAlign: "left", fontFamily: "inherit",
              }}>
                <div style={{ width: 48, height: 48, borderRadius: 8, background: t.logoBg, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 }}>
                  <img src={t.logo} alt={t.name} style={{ width: "100%", height: "100%", objectFit: "contain" }} />
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{t.name}</div>
                  <div style={{ fontSize: 12, color: "#8B93A1" }}>{t.tagline}</div>
                </div>
                <ChevronRight size={16} color="#5C6470" />
              </button>
            ))}
          </div>
        )}

        {picked && !busy && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, padding: 12, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10 }}>
              <div style={{ width: 36, height: 36, borderRadius: 6, background: picked.logoBg, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 }}>
                <img src={picked.logo} alt={picked.name} style={{ width: "100%", height: "100%", objectFit: "contain" }} />
              </div>
              <div style={{ fontSize: 13.5, fontWeight: 600 }}>{picked.name}</div>
            </div>

            {mode === "create" && (
              <div style={{ fontSize: 12.5, color: "#8B93A1" }}>First time for this team — choose an access code that will be required for everyone from now on.</div>
            )}

            <input autoFocus type="password" placeholder={mode === "create" ? "New access code" : "Access code"}
              value={code} onChange={e => setCode(e.target.value)} style={inputStyle} />
            {mode === "create" && (
              <input type="password" placeholder="Confirm the code" value={confirmCode} onChange={e => setConfirmCode(e.target.value)} style={inputStyle} />
            )}
            {error && <div style={{ color: RED, fontSize: 13 }}>{error}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { setPicked(null); setError(""); }} style={btnSecondary}>Back</button>
              <button onClick={mode === "create" ? submitCreate : submitEnter} style={btnPrimary}>
                {mode === "create" ? "Create access" : "Enter"}
              </button>
            </div>
          </div>
        )}

        {!picked && (
          <button onClick={() => setShowAdmin(true)} style={{ display: "block", margin: "18px auto 0", background: "none", border: "none", color: "#4A5361", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>
            Admin
          </button>
        )}
      </div>
    </div>
  );
}

// Accès direct à une clé complète, sans le préfixage automatique de storeGet/storeSet —
// nécessaire pour le panneau Admin, qui doit lire/écrire les données de N'IMPORTE QUELLE
// équipe sans "entrer" dedans (donc sans changer TEAM_PREFIX). Même principe que la
// sauvegarde/restauration.
async function rawGet(fullKey) {
  await supabaseInit;
  try {
    if (window.storage) { const r = await window.storage.get(fullKey, true); return r ? JSON.parse(r.value) : null; }
    if (supabase) {
      const { data, error } = await supabase.from("app_storage").select("value").eq("key", fullKey).maybeSingle();
      if (error) console.error("[storage] Supabase get error on", fullKey, ":", error);
      return data ? data.value : null;
    }
  } catch (e) { console.error("[storage] rawGet failed on", fullKey, ":", e); }
  return memoryStore[fullKey] ?? null;
}
async function rawSet(fullKey, value) {
  await supabaseInit;
  memoryStore[fullKey] = value;
  if (window.storage) { await window.storage.set(fullKey, JSON.stringify(value), true); return; }
  if (supabase) {
    const { error } = await supabase.from("app_storage").upsert({ key: fullKey, value });
    if (error) { console.error("[storage] Supabase rawSet FAILED on", fullKey, ":", error); throw error; }
  }
}
async function rawDelete(fullKey) {
  await supabaseInit;
  try {
    if (window.storage) await window.storage.delete(fullKey, true);
    else if (supabase) await supabase.from("app_storage").delete().eq("key", fullKey);
  } catch (e) { console.error("[storage] rawDelete failed on", fullKey, ":", e); }
  delete memoryStore[fullKey];
}

// ---------------------------------------------------------------------------
// Deletion requests — coaches don't delete directly; they file a request that
// the Admin must approve or reject from the Admin panel.
// ---------------------------------------------------------------------------

async function requestDeletion(teamId, teamName, type, label, meta) {
  const pending = (await rawGet("pending_deletions")) || [];
  pending.push({ id: uid(), teamId, teamName, type, label, meta, requestedAt: new Date().toISOString() });
  await rawSet("pending_deletions", pending);
}

async function approveDeletion(req) {
  const p = "team_" + req.teamId + ":";
  if (req.type === "match") {
    const idx = (await rawGet(p + "match_index")) || [];
    await rawSet(p + "match_index", idx.filter(m => m.id !== req.meta.id));
    await rawDelete(p + "match:" + req.meta.id);
    await rawDelete(p + "match_file:" + req.meta.id);
  } else if (req.type === "boxscore") {
    const idx = (await rawGet(p + "boxscore_index")) || [];
    await rawSet(p + "boxscore_index", idx.filter(m => m.id !== req.meta.id));
    await rawDelete(p + "boxscore:" + req.meta.id);
    await rawDelete(p + "boxscore_file:" + req.meta.id);
  } else if (req.type === "player") {
    const roster = (await rawGet(p + "roster")) || [];
    await rawSet(p + "roster", roster.filter(pl => pl.id !== req.meta.id));
  } else if (req.type === "wellness") {
    const key = p + "wellness:" + req.meta.playerName;
    const entries = (await rawGet(key)) || [];
    await rawSet(key, entries.filter(e => e.id !== req.meta.entryId));
  } else if (req.type === "resource") {
    const key = p + "team_resources";
    const resources = (await rawGet(key)) || [];
    await rawSet(key, resources.filter(r => r.id !== req.meta.id));
  } else if (req.type === "planning_event") {
    const key = p + "planning_events";
    const evs = (await rawGet(key)) || [];
    await rawSet(key, evs.filter(e => e.id !== req.meta.id));
  }
  const pending = (await rawGet("pending_deletions")) || [];
  await rawSet("pending_deletions", pending.filter(r => r.id !== req.id));
}

async function rejectDeletion(req) {
  const pending = (await rawGet("pending_deletions")) || [];
  await rawSet("pending_deletions", pending.filter(r => r.id !== req.id));
}

// ---------------------------------------------------------------------------
// Admin — gestion des équipes et des codes d'accès (joueurs/coachs), protégée par
// un mot de passe séparé des codes d'équipe.
// ---------------------------------------------------------------------------

function AdminGate({ onClose, teams, onTeamsChange }) {
  const [unlocked, setUnlocked] = useState(false);
  const [mode, setMode] = useState(null); // "enter" | "create"
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");

  useEffect(() => { rawGet("admin_password").then(p => setMode(p ? "enter" : "create")); }, []);

  async function submit() {
    if (mode === "create") {
      if (password.length < 4) { setError("Choose a password of at least 4 characters."); return; }
      if (password !== confirmPassword) { setError("Passwords don't match."); return; }
      await rawSet("admin_password", simpleHash(password));
      setUnlocked(true);
    } else {
      const stored = await rawGet("admin_password");
      if (simpleHash(password) !== stored) { setError("Incorrect password."); return; }
      setUnlocked(true);
    }
  }

  if (unlocked) return <AdminPanel onClose={onClose} teams={teams} onTeamsChange={onTeamsChange} />;

  return (
    <div style={{ minHeight: "100vh", background: INK, color: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 20 }}>
      <div style={{ width: 380, maxWidth: "100%" }}>
        <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontSize: 12, letterSpacing: "0.18em", marginBottom: 6 }}>ADMIN</div>
        <h1 style={{ fontSize: 22, fontWeight: 800, margin: "0 0 18px" }}>{mode === "create" ? "Set an admin password" : "Admin password"}</h1>
        {mode === "create" && <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 14 }}>First time here — this password will be required for admin access from now on.</div>}
        <input autoFocus type="password" placeholder="Password" value={password} onChange={e => setPassword(e.target.value)} style={{ ...inputStyle, marginBottom: 10 }} />
        {mode === "create" && <input type="password" placeholder="Confirm password" value={confirmPassword} onChange={e => setConfirmPassword(e.target.value)} style={{ ...inputStyle, marginBottom: 10 }} />}
        {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onClose} style={btnSecondary}>Back</button>
          {mode && <button onClick={submit} style={btnPrimary}>{mode === "create" ? "Set password" : "Unlock"}</button>}
        </div>
      </div>
    </div>
  );
}

function PendingDeletionsBox() {
  const [pending, setPending] = useState(null);
  const [busy, setBusy] = useState(null);

  useEffect(() => { load(); }, []);
  async function load() { setPending((await rawGet("pending_deletions")) || []); }

  async function approve(req) { setBusy(req.id); await approveDeletion(req); await load(); setBusy(null); }
  async function reject(req) { setBusy(req.id); await rejectDeletion(req); await load(); setBusy(null); }

  if (pending === null || pending.length === 0) return null;

  return (
    <div style={{ background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 12, padding: 18, marginBottom: 24 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <AlertTriangle size={15} color={AMBER} />
        <div style={{ fontSize: 14, fontWeight: 700 }}>Pending deletions ({pending.length})</div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {pending.map(req => (
          <div key={req.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, padding: "10px 12px", flexWrap: "wrap", gap: 8 }}>
            <div style={{ fontSize: 12.5 }}>
              <b>{req.label}</b> <span style={{ color: "#5C6470" }}>· {req.type} · {req.teamName}</span>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button disabled={busy === req.id} onClick={() => approve(req)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11.5, padding: "6px 10px", cursor: "pointer" }}>Approve delete</button>
              <button disabled={busy === req.id} onClick={() => reject(req)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11.5, padding: "6px 10px", cursor: "pointer" }}>Reject</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function AdminPanel({ onClose, teams, onTeamsChange }) {
  const [newTeamName, setNewTeamName] = useState("");
  const [newTeamLogo, setNewTeamLogo] = useState("");
  const [confirmDeleteTeam, setConfirmDeleteTeam] = useState(null);
  const [expandedTeam, setExpandedTeam] = useState(null);
  const [busy, setBusy] = useState(false);
  const newLogoRef = useRef();

  async function handleNewLogo(e) {
    const file = e.target.files[0];
    if (!file) return;
    setNewTeamLogo(await fileToResizedDataURL(file, 300, 0.9));
  }

  async function addTeam() {
    const name = newTeamName.trim();
    if (!name) return;
    const id = name.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || uid();
    const palette = [AMBER, TEAL, "#7C9CF2", "#C97BE0", "#E15A4E"];
    const next = [...teams, { id, name, tagline: "", logo: newTeamLogo, logoBg: palette[teams.length % palette.length] }];
    setBusy(true);
    await saveTeams(next);
    onTeamsChange(next);
    setNewTeamName(""); setNewTeamLogo(""); setBusy(false);
  }

  async function removeTeam(id) {
    const next = teams.filter(t => t.id !== id);
    setBusy(true);
    await saveTeams(next);
    onTeamsChange(next);
    setConfirmDeleteTeam(null); setBusy(false);
  }

  async function updateTeam(id, patch) {
    const next = teams.map(t => t.id === id ? { ...t, ...patch } : t);
    setBusy(true);
    await saveTeams(next);
    onTeamsChange(next);
    setBusy(false);
  }

  return (
    <div style={{ minHeight: "100vh", background: INK, color: PAPER, fontFamily: "'Inter', sans-serif", padding: "28px 20px 60px" }}>
      <div style={{ maxWidth: 760, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <div>
            <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontSize: 12, letterSpacing: "0.18em", marginBottom: 6 }}>ADMIN</div>
            <h1 style={{ fontSize: 22, fontWeight: 800, margin: 0 }}>Teams & access</h1>
          </div>
          <button onClick={onClose} style={btnSecondary}>Close</button>
        </div>

        <PendingDeletionsBox />

        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 24 }}>
          <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>Add a team</div>
          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            <button type="button" onClick={() => newLogoRef.current && newLogoRef.current.click()} style={{ width: 46, height: 46, borderRadius: 8, background: PANEL2, border: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0, cursor: "pointer", padding: 0 }}>
              {newTeamLogo ? <img src={newTeamLogo} alt="" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> : <Camera size={16} color="#5C6470" />}
            </button>
            <input ref={newLogoRef} type="file" accept="image/*" onChange={handleNewLogo} style={{ display: "none" }} />
            <input value={newTeamName} onChange={e => setNewTeamName(e.target.value)} placeholder="Team name" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", flex: 1 }} />
            <button disabled={busy || !newTeamName.trim()} onClick={addTeam} style={{ ...btnPrimary, width: "auto", padding: "10px 18px" }}>Add</button>
          </div>
        </div>

        {teams.map(team => (
          <TeamAdminCard
            key={team.id}
            team={team}
            expanded={expandedTeam === team.id}
            onToggle={() => setExpandedTeam(expandedTeam === team.id ? null : team.id)}
            confirmDelete={confirmDeleteTeam === team.id}
            onAskDelete={() => setConfirmDeleteTeam(team.id)}
            onCancelDelete={() => setConfirmDeleteTeam(null)}
            onConfirmDelete={() => removeTeam(team.id)}
            onUpdateLogo={(logo) => updateTeam(team.id, { logo })}
          />
        ))}
      </div>
    </div>
  );
}

function TeamAdminCard({ team, expanded, onToggle, confirmDelete, onAskDelete, onCancelDelete, onConfirmDelete, onUpdateLogo }) {
  const [users, setUsers] = useState(null);
  const [newCode, setNewCode] = useState("");
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);
  const [visibility, setVisibility] = useState(null);
  const logoRef = useRef();

  useEffect(() => {
    if (expanded) {
      rawGet("team_" + team.id + ":app_users").then(u => setUsers(u || {}));
      rawGet("team_" + team.id + ":visibility_config").then(v => setVisibility(v ? { ...DEFAULT_VISIBILITY, ...v, tabs: { ...DEFAULT_VISIBILITY.tabs, ...(v.tabs || {}) }, playerDetail: { ...DEFAULT_VISIBILITY.playerDetail, ...(v.playerDetail || {}) }, team: { ...DEFAULT_VISIBILITY.team, ...(v.team || {}) } } : DEFAULT_VISIBILITY));
    }
  }, [expanded]);

  async function saveVisibility(next) {
    setVisibility(next);
    await rawSet("team_" + team.id + ":visibility_config", next);
  }
  function toggleTab(key) { saveVisibility({ ...visibility, tabs: { ...visibility.tabs, [key]: !visibility.tabs[key] } }); }
  function togglePlayerDetail(key) { saveVisibility({ ...visibility, playerDetail: { ...visibility.playerDetail, [key]: !visibility.playerDetail[key] } }); }
  function toggleTeamSection(key) { saveVisibility({ ...visibility, team: { ...visibility.team, [key]: !visibility.team[key] } }); }
  function toggleWellnessCharts() { saveVisibility({ ...visibility, wellnessCharts: !visibility.wellnessCharts }); }

  async function handleLogoChange(e) {
    const file = e.target.files[0];
    if (!file) return;
    const dataUrl = await fileToResizedDataURL(file, 300, 0.9);
    onUpdateLogo(dataUrl);
  }

  async function resetCode() {
    if (newCode.length < 4) { setStatus("Code must be at least 4 characters."); return; }
    setBusy(true);
    await rawSet("team_access:" + team.id, simpleHash(newCode));
    setStatus("Access code updated.");
    setNewCode(""); setBusy(false);
  }

  async function removeUser(name) {
    const next = { ...users };
    delete next[name];
    setBusy(true);
    await rawSet("team_" + team.id + ":app_users", next);
    setUsers(next); setBusy(false);
  }

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, marginBottom: 14, overflow: "hidden" }}>
      <div style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", padding: 16 }}>
        <button onClick={onToggle} style={{ flex: 1, display: "flex", alignItems: "center", gap: 12, background: "none", border: "none", cursor: "pointer", color: PAPER, fontFamily: "inherit", padding: 0, textAlign: "left" }}>
          <div style={{ position: "relative", width: 34, height: 34, flexShrink: 0 }}>
            <div style={{ width: 34, height: 34, borderRadius: 7, background: team.logoBg, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
              {team.logo && <img src={team.logo} alt="" style={{ width: "100%", height: "100%", objectFit: "contain" }} />}
            </div>
          </div>
          <div style={{ fontWeight: 700, fontSize: 14.5 }}>{team.name}</div>
        </button>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button type="button" onClick={() => logoRef.current && logoRef.current.click()} title="Change logo" style={{ fontSize: 11.5, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>Change logo</button>
          <input ref={logoRef} type="file" accept="image/*" onChange={handleLogoChange} style={{ display: "none" }} />
          <button onClick={onToggle} style={{ background: "none", border: "none", cursor: "pointer", display: "flex" }}><ChevronRight size={16} color="#5C6470" style={{ transform: expanded ? "rotate(90deg)" : "none" }} /></button>
        </div>
      </div>

      {expanded && (
        <div style={{ padding: "0 16px 16px" }}>
          <div style={{ height: 1, background: LINE, marginBottom: 16 }} />

          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Reset team access code</div>
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <input type="password" value={newCode} onChange={e => setNewCode(e.target.value)} placeholder="New access code" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", flex: 1, maxWidth: 240 }} />
            <button disabled={busy} onClick={resetCode} style={btnSecondary}>Set</button>
          </div>
          {status && <div style={{ fontSize: 12, color: TEAL, marginBottom: 14 }}>{status}</div>}

          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Visible to players</div>
          {visibility && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 16 }}>
              {[["players", "Players tab"], ["team", "Team tab"], ["scouting", "Scouting tab"], ["planning", "Planning tab"]].map(([key, label]) => (
                <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={visibility.tabs[key]} onChange={() => toggleTab(key)} /> {label}
                </label>
              ))}
              <div style={{ height: 1, background: LINE, margin: "4px 0" }} />
              {[["standings", "Team — Standings"], ["teamPlay", "Team — Team Play"], ["advanced", "Team — Advanced"], ["resources", "Team — Resources"]].map(([key, label]) => (
                <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={visibility.team[key]} onChange={() => toggleTeamSection(key)} /> {label}
                </label>
              ))}
              <div style={{ height: 1, background: LINE, margin: "4px 0" }} />
              {[["stats", "Match Stats sub-tab"], ["objectives", "Objectives sub-tab"], ["training", "Training sub-tab"], ["mental", "Mental evaluation sub-tab"], ["wellness", "Wellness sub-tab"], ["role", "Role sub-tab"], ["meetings", "Meetings sub-tab"]].map(([key, label]) => (
                <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={visibility.playerDetail[key]} onChange={() => togglePlayerDetail(key)} /> {label}
                </label>
              ))}
              <div style={{ height: 1, background: LINE, margin: "4px 0" }} />
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, cursor: "pointer" }}>
                <input type="checkbox" checked={visibility.wellnessCharts} onChange={toggleWellnessCharts} /> Wellness charts (players can see their own trends)
              </label>
            </div>
          )}

          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Players & coaches with an account</div>
          {users === null ? (
            <div style={{ fontSize: 12.5, color: "#5C6470" }}>Loading…</div>
          ) : Object.keys(users).length === 0 ? (
            <div style={{ fontSize: 12.5, color: "#5C6470" }}>No one has logged in yet for this team.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 16 }}>
              {Object.entries(users).map(([name, u]) => (
                <div key={name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, padding: "8px 12px" }}>
                  <div style={{ fontSize: 13 }}>{name} <span style={{ fontSize: 11, color: "#5C6470" }}>· {u.role === "coach" ? "Staff" : "Player"}</span></div>
                  <button onClick={() => removeUser(name)} title="Revoke access" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={14} /></button>
                </div>
              ))}
            </div>
          )}

          <div style={{ height: 1, background: LINE, margin: "6px 0 14px" }} />
          {!confirmDelete ? (
            <button onClick={onAskDelete} style={{ fontSize: 12.5, color: RED, background: "none", border: "none", cursor: "pointer" }}>Delete this team</button>
          ) : (
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <span style={{ fontSize: 12.5, color: RED }}>Delete "{team.name}"? This only removes it from the list — its data stays in storage.</span>
              <button onClick={onConfirmDelete} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 12, padding: "5px 10px", cursor: "pointer" }}>Yes, delete</button>
              <button onClick={onCancelDelete} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 12, padding: "5px 10px", cursor: "pointer" }}>Cancel</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function LoginScreen({ roster, onLogin }) {
  const [step, setStep] = useState("pick"); // pick | pin | newpin
  const [name, setName] = useState(""); // identité de stockage (prénom pour un joueur, libellé pour le staff)
  const [displayName, setDisplayName] = useState(""); // ce qui s'affiche à l'écran
  const [pin, setPin] = useState("");
  const [confirmPin, setConfirmPin] = useState("");
  const [error, setError] = useState("");
  const [existingUser, setExistingUser] = useState(null);

  const staffOptions = ["Head coach", "Assistant coach", "Analyst"];
  // Le nom complet sert désormais d'identité de stockage ET d'affichage — c'est ce même nom
  // complet qui identifie ce joueur partout ailleurs dans l'appli (entraînement, objectifs,
  // Wellness, évaluation mentale...), pour éviter toute confusion entre deux joueurs qui
  // partageraient le même prénom.
  const nameOptions = [
    ...staffOptions.map(s => ({ key: s, label: s })),
    ...roster.map(p => ({ key: p.name, label: p.name })),
  ];

  async function chooseName(opt) {
    setName(opt.key);
    setDisplayName(opt.label);
    setError("");
    const users = (await storeGet("app_users")) || {};
    if (users[opt.key]) { setExistingUser(users[opt.key]); setStep("pin"); }
    else { setStep("newpin"); }
  }

  async function submitPin() {
    if (pin.length !== 4) { setError("The code must be 4 digits."); return; }
    const users = (await storeGet("app_users")) || {};
    const u = users[name];
    if (simpleHash(pin) !== u.pin) { setError("Incorrect code."); return; }
    const match = roster.find(p => p.name === name);
    onLogin({ name, role: u.role, id: u.role === "player" ? match?.id : undefined });
  }

  async function submitNewPin() {
    if (pin.length !== 4) { setError("Choose a 4-digit code."); return; }
    if (pin !== confirmPin) { setError("The two codes don't match."); return; }
    const users = (await storeGet("app_users")) || {};
    const role = staffOptions.includes(name) ? "coach" : "player";
    users[name] = { pin: simpleHash(pin), role };
    await storeSet("app_users", users);
    const match = roster.find(p => p.name === name);
    onLogin({ name, role, id: role === "player" ? match?.id : undefined });
  }

  return (
    <div style={{ minHeight: "100vh", background: INK, color: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 20 }}>
      <div style={{ width: 380, maxWidth: "100%" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center", marginBottom: 18 }}>
          <img src="/icon-512.png" alt="" onError={e => { e.target.style.display = "none"; }} style={{ width: 56, height: 56, borderRadius: 12, marginBottom: 10 }} />
          <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontSize: 12, fontWeight: 700, letterSpacing: "0.14em" }}>HOOPTRACK BASKETBALL</div>
        </div>
        <h1 style={{ fontSize: 26, fontWeight: 800, margin: "0 0 22px", letterSpacing: "-0.01em", textAlign: "center" }}>
          {step === "pick" ? "Who are you?" : step === "pin" ? `${displayName}'s code` : `New code — ${displayName}`}
        </h1>

        {step === "pick" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 420, overflowY: "auto" }}>
            {nameOptions.map(opt => (
              <button key={opt.key} onClick={() => chooseName(opt)} style={btnRow}>
                <span>{opt.label}</span><ChevronRight size={16} color="#5C6470" />
              </button>
            ))}
          </div>
        )}

        {(step === "pin" || step === "newpin") && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input autoFocus type="password" inputMode="numeric" maxLength={4} placeholder="4-digit code"
              value={pin} onChange={e => setPin(e.target.value.replace(/\D/g, ""))} style={inputStyle} />
            {step === "newpin" && (
              <input type="password" inputMode="numeric" maxLength={4} placeholder="Confirm the code"
                value={confirmPin} onChange={e => setConfirmPin(e.target.value.replace(/\D/g, ""))} style={inputStyle} />
            )}
            {error && <div style={{ color: RED, fontSize: 13 }}>{error}</div>}
            <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
              <button onClick={() => { setStep("pick"); setError(""); setPin(""); setConfirmPin(""); }} style={btnSecondary}>Back</button>
              <button onClick={step === "pin" ? submitPin : submitNewPin} style={btnPrimary}>Enter</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const btnRow = { display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", padding: "12px 14px", background: PANEL, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 14, cursor: "pointer", textAlign: "left", fontFamily: "inherit" };
const inputStyle = { width: "100%", padding: "12px 14px", background: PANEL, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 16, letterSpacing: "0.2em", fontFamily: "ui-monospace, monospace", boxSizing: "border-box" };
const btnPrimary = { flex: 1, padding: "11px 14px", background: AMBER, border: "none", borderRadius: 8, color: "#1A1300", fontWeight: 700, cursor: "pointer", fontSize: 14 };
const btnSecondary = { padding: "11px 14px", background: "transparent", border: `1px solid ${LINE}`, borderRadius: 8, color: "#8B93A1", cursor: "pointer", fontSize: 14 };

// ---------------------------------------------------------------------------
// Top bar
// ---------------------------------------------------------------------------

function TopBar({ session, onSwitchAccount, onSwitchTeam, team, isCoach }) {
  return (
    <div style={{ borderBottom: `1px solid ${LINE}`, background: "#0B0D11" }}>
      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "14px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <button onClick={onSwitchTeam} title="Switch team" style={{ display: "flex", alignItems: "center", gap: 8, background: "none", border: "none", cursor: "pointer", padding: 0 }}>
          <div style={{ width: 24, height: 24, borderRadius: 5, background: team.logoBg, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 }}>
            <img src={team.logo} alt={team.name} style={{ width: "100%", height: "100%", objectFit: "contain" }} />
          </div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontWeight: 700, fontSize: 12, letterSpacing: "0.06em", textAlign: "left" }}>{team.name}</div>
        </button>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ fontSize: 12.5, color: "#8B93A1", display: "flex", alignItems: "center", gap: 6 }}>
            <ShieldCheck size={13} color={isCoach ? TEAL : "#8B93A1"} />
            {session.name} <span style={{ color: "#4A5361" }}>· {isCoach ? "Staff" : "Player"}</span>
          </div>
          <button onClick={onSwitchAccount} title="Switch account" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><LogOut size={16} /></button>
        </div>
      </div>
    </div>
  );
}

// Barre d'onglets fixée en bas de l'écran, façon applications mobiles (icône au-dessus
// du libellé, appui direct) — remplace l'ancienne navigation horizontale en haut.
// ---------------------------------------------------------------------------
// Home — tableau de bord condensé affiché en premier après connexion. Réutilise
// volontairement les mêmes composants/hooks que les pages détaillées (pas de
// nouvelle logique de calcul) pour rester cohérent et facile à maintenir.
// ---------------------------------------------------------------------------

// Petit bandeau de section cliquable — même apparence que SectionTitle mais qui amène
// directement vers la page complète correspondante quand on clique dessus.
function HomeSectionLink({ eyebrow, title, onClick }) {
  return (
    <button onClick={onClick} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", padding: 0, marginBottom: 12, cursor: "pointer", fontFamily: "inherit", textAlign: "left", color: PAPER }}>
      <SectionTitle eyebrow={eyebrow} title={title} />
      <ChevronRight size={16} color="#5C6470" />
    </button>
  );
}

function HomeTab({ session, isCoach, playerName, allPlays, roster, matchFilter, team, goToPlayer, goToScouting, goToTeam, goToPlanning, visibility }) {
  const box = useBoxScore(playerName, matchFilter);
  const advanced = useTeamAdvancedStats(matchFilter);
  const objectives = useObjectives(playerName || "");
  const [trainings, setTrainings] = useState([]);
  const [mentalEntries, setMentalEntries] = useState([]);
  const [wellnessEntries, setWellnessEntries] = useState([]);
  const [wSlot, setWSlot] = useState("wake");
  const [wPhysical, setWPhysical] = useState(3);
  const [wMental, setWMental] = useState(3);
  const [wBusy, setWBusy] = useState(false);
  const [wStatus, setWStatus] = useState("");
  const scouting = useScoutingTeams();
  const [nextGame, setNextGame] = useState("");
  const [role, setRole] = useState(null);
  const [resources, setResources] = useState([]);
  const [todayEvents, setTodayEvents] = useState([]);
  const [homeMessage, setHomeMessage] = useState("");
  const [editingMessage, setEditingMessage] = useState(false);
  const [messageDraft, setMessageDraft] = useState("");
  const [messageBusy, setMessageBusy] = useState(false);

  useEffect(() => { storeGet("home_message").then(m => setHomeMessage(m || "")); }, []);
  async function saveHomeMessage() {
    setMessageBusy(true);
    try { await storeSet("home_message", messageDraft.trim()); setHomeMessage(messageDraft.trim()); }
    finally { setMessageBusy(false); setEditingMessage(false); }
  }
  async function clearHomeMessage() {
    setMessageBusy(true);
    try { await storeSet("home_message", ""); setHomeMessage(""); }
    finally { setMessageBusy(false); }
  }

  useEffect(() => { storeGet("team_resources").then(r => setResources(((r || []).sort((a, b) => b.addedAt.localeCompare(a.addedAt))).slice(0, 2))); }, []);
  useEffect(() => {
    storeGet("planning_events").then(evs => {
      const todayKey = toDateKey(new Date());
      const me = playerName ? roster.find(p => p.name === playerName) : null;
      const relevant = (evs || []).filter(e => e.date === todayKey && (!e.playerIds || e.playerIds.length === 0 || (me && e.playerIds.includes(me.id))));
      setTodayEvents(relevant.sort((a, b) => a.startTime.localeCompare(b.startTime)));
    });
  }, []);

  useEffect(() => {
    if (playerName) {
      storeGet("training:" + playerName).then(t => setTrainings((t || []).slice(0, 3)));
      storeGet("mental:" + playerName).then(m => setMentalEntries(m || []));
      storeGet("wellness:" + playerName).then(w => setWellnessEntries(w || []));
      storeGet("role:" + playerName).then(r => setRole(r || null));
    }
  }, [playerName]);

  useEffect(() => { storeGet("next_game_team").then(n => setNextGame(n || "")); }, []);

  async function saveNextGame(name) {
    setNextGame(name);
    await storeSet("next_game_team", name);
  }

  async function submitWellness() {
    setWBusy(true); setWStatus("");
    const today = todayLocal();
    try {
      // BUG RÉEL CORRIGÉ : se baser sur l'état React local (wellnessEntries) au lieu de relire
      // la donnée la plus récente juste avant d'écrire pouvait écraser une réponse qui venait
      // juste d'être enregistrée si le joueur répondait à deux créneaux coup sur coup, avant que
      // l'état local n'ait eu le temps de se mettre à jour — certaines réponses "disparaissaient".
      const latest = (await storeGet("wellness:" + playerName)) || [];
      const existingIdx = latest.findIndex(e => e.date === today && e.slot === wSlot);
      const entry = { id: existingIdx >= 0 ? latest[existingIdx].id : uid(), date: today, slot: wSlot, physical: wPhysical, mental: wMental };
      const next = existingIdx >= 0 ? latest.map((e, i) => i === existingIdx ? entry : e) : [entry, ...latest];
      await storeSet("wellness:" + playerName, next);
      setWellnessEntries(next);
      setWStatus("Saved — thanks!");
    } catch (e) {
      // Ne JAMAIS afficher "Saved" si l'écriture a réellement échoué (ex. mauvaise connexion) —
      // bug réel constaté où un joueur voyait une confirmation alors que la réponse n'avait
      // jamais atteint la base de données.
      setWStatus("Not saved — check your connection and try again.");
    }
    setWBusy(false);
  }

  const plays = playerName ? allPlays.filter(p => p.player === playerName) : [];
  const off = plays.filter(isOffense);
  const def = plays.filter(isDefense);

  const teamOff = allPlays.filter(isOffense);
  const teamDef = allPlays.filter(isDefense);

  const avg = (key) => {
    const vals = advanced.perMatch.map(m => m[key]).filter(v => v !== null && v !== undefined);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  };
  const { efg, tovPct, orebPct } = computeWeightedTeamPercentages(advanced.perMatch);
  const ortg = avg("ortg"), drtg = avg("drtg"), ftRate = avg("ftRate"), oreb = avg("oreb");

  const today = todayLocal();
  const answeredToday = new Set(wellnessEntries.filter(e => e.date === today).map(e => e.slot));
  const mentalAvg = mentalEntries.length
    ? mentalEntries.reduce((s, e) => s + (Object.values(e.ratings || {}).reduce((a, b) => a + b, 0) / (Object.values(e.ratings || {}).length || 1)), 0) / mentalEntries.length
    : null;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, marginBottom: 24, flexWrap: "wrap" }}>
        <div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontSize: 12, letterSpacing: "0.1em", marginBottom: 4 }}>WELCOME</div>
          <h1 style={{ fontSize: 22, fontWeight: 800, margin: 0 }}>{session.name}</h1>
        </div>

        <div style={{ maxWidth: 280, width: "100%", flexShrink: 0 }}>
          {isCoach ? (
            editingMessage ? (
              <div style={{ background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 10, padding: 12 }}>
                <textarea value={messageDraft} onChange={e => setMessageDraft(e.target.value)} rows={3} placeholder="Message shown to players on their Home screen…"
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical", fontSize: 13, marginBottom: 8 }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button disabled={messageBusy} onClick={saveHomeMessage} style={{ ...btnPrimary, width: "auto", padding: "6px 14px", fontSize: 12.5 }}>{messageBusy ? "…" : "Save"}</button>
                  <button onClick={() => setEditingMessage(false)} style={{ ...btnSecondary, padding: "6px 14px", fontSize: 12.5 }}>Cancel</button>
                </div>
              </div>
            ) : homeMessage ? (
              <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 12 }}>
                <div style={{ fontSize: 12.5, color: PAPER, lineHeight: 1.5, whiteSpace: "pre-wrap", marginBottom: 8 }}>{homeMessage}</div>
                <div style={{ display: "flex", gap: 12 }}>
                  <button onClick={() => { setMessageDraft(homeMessage); setEditingMessage(true); }} style={{ fontSize: 11.5, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>Edit</button>
                  <button onClick={clearHomeMessage} style={{ fontSize: 11.5, color: "#5C6470", background: "none", border: "none", cursor: "pointer" }}>Remove</button>
                </div>
              </div>
            ) : (
              <button onClick={() => { setMessageDraft(""); setEditingMessage(true); }} style={{ fontSize: 12, color: "#5C6470", background: "none", border: `1px dashed ${LINE}`, borderRadius: 10, padding: "10px 12px", cursor: "pointer", width: "100%", textAlign: "left" }}>
                + Write a message for players' Home screen
              </button>
            )
          ) : (
            homeMessage && (
              <div style={{ background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 10, padding: 12 }}>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: AMBER, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>Message from the coach</div>
                <div style={{ fontSize: 12.5, color: PAPER, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{homeMessage}</div>
              </div>
            )
          )}
        </div>
      </div>

      {(isCoach || (nextGame && (visibility || DEFAULT_VISIBILITY).tabs.scouting)) && (
        <div style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Next game</div>
          {isCoach && (
            <select value={nextGame} onChange={e => saveNextGame(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320, marginBottom: nextGame ? 10 : 0 }}>
              <option value="">— Select the scouted opponent —</option>
              {Object.keys(scouting.teams).map(n => <option key={n} value={n}>{n}</option>)}
            </select>
          )}
          {nextGame && scouting.teams[nextGame] && (
            <button onClick={() => goToScouting("rapport", nextGame)} style={{ display: "flex", alignItems: "center", gap: 14, width: "100%", maxWidth: 320, padding: "14px 16px", background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 10, color: PAPER, cursor: "pointer", textAlign: "left", fontFamily: "inherit" }}>
              <div style={{ width: 44, height: 44, borderRadius: 8, background: PANEL2, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 }}>
                {scouting.teams[nextGame].logo ? <img src={scouting.teams[nextGame].logo} alt="" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> : <span style={{ fontSize: 18, fontWeight: 800, color: AMBER }}>{nextGame[0]}</span>}
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 10.5, color: AMBER, fontWeight: 700, letterSpacing: "0.08em" }}>NEXT GAME</div>
                <div style={{ fontWeight: 700, fontSize: 15 }}>{nextGame}</div>
              </div>
              <ChevronRight size={16} color="#5C6470" />
            </button>
          )}
        </div>
      )}

      {playerName && (() => { const pd = (visibility || DEFAULT_VISIBILITY).playerDetail || DEFAULT_VISIBILITY.playerDetail; return pd.role; })() && role && (role.name || role.image) && (
        <div style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Role</div>
          <button onClick={() => goToPlayer("role")} style={{ display: "block", width: "100%", maxWidth: 320, padding: 0, background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, color: PAPER, cursor: "pointer", textAlign: "left", fontFamily: "inherit", overflow: "hidden" }}>
            {role.image ? (
              <img src={role.image} alt="" style={{ width: "100%", maxHeight: 120, objectFit: "contain", background: PANEL2, display: "block" }} />
            ) : (
              <div style={{ width: "100%", height: 80, background: PANEL2, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <ShieldCheck size={28} color={AMBER} />
              </div>
            )}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px" }}>
              <div style={{ fontWeight: 700, fontSize: 15 }}>{role.name || "Role"}</div>
              <ChevronRight size={16} color="#5C6470" />
            </div>
          </button>
        </div>
      )}

      {resources.length > 0 && (
        <div style={{ marginBottom: 26 }}>
          <HomeSectionLink eyebrow="Latest" title="Resources" onClick={goToTeam} />
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {resources.map(r => (
              <a key={r.id} href={r.url} target="_blank" rel="noopener noreferrer" style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 14px", background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, color: PAPER, textDecoration: "none" }}>
                {r.type === "video" ? <Video size={16} color={AMBER} /> : <LinkIcon size={16} color={TEAL} />}
                <span style={{ fontSize: 13.5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.title}</span>
              </a>
            ))}
          </div>
        </div>
      )}

      {playerName && (() => {
        const pd = (visibility || DEFAULT_VISIBILITY).playerDetail || DEFAULT_VISIBILITY.playerDetail;
        return (
        <>
          {pd.wellness && (
            <>
              <HomeSectionLink eyebrow="3x a day" title="Wellness" onClick={() => goToPlayer("wellness")} />
              <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 26 }}>
                <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 12 }}>
                  Today: {WELLNESS_SLOTS.map(s => answeredToday.has(s.key) ? `✓ ${s.label}` : s.label).join(" · ")}
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 12 }}>
                  <div style={{ width: 170 }}>
                    <label style={labelStyle}>When</label>
                    <select value={wSlot} onChange={e => setWSlot(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
                      {WELLNESS_SLOTS.map(s => <option key={s.key} value={s.key}>{s.label}</option>)}
                    </select>
                  </div>
                  <div style={{ width: 130 }}>
                    <label style={labelStyle}>Physical</label>
                    <select value={wPhysical} onChange={e => setWPhysical(Number(e.target.value))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: ratingColor(wPhysical), fontWeight: 700 }}>
                      {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  </div>
                  <div style={{ width: 130 }}>
                    <label style={labelStyle}>Mental</label>
                    <select value={wMental} onChange={e => setWMental(Number(e.target.value))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: ratingColor(wMental), fontWeight: 700 }}>
                      {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  </div>
                  <button disabled={wBusy} onClick={submitWellness} style={{ ...btnPrimary, width: "auto", padding: "10px 18px" }}>{wBusy ? "…" : "Submit"}</button>
                </div>
                {wStatus && <div style={{ fontSize: 12, color: TEAL }}>{wStatus}</div>}
              </div>
            </>
          )}

          {(visibility || DEFAULT_VISIBILITY).tabs.planning && (
            <>
              <HomeSectionLink eyebrow={new Date().toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" })} title="Today's schedule" onClick={goToPlanning} />
              {todayEvents.length === 0 ? <EmptyState text="Nothing scheduled today." /> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 26 }}>
                  {todayEvents.map(ev => (
                    <div key={ev.id} style={{ display: "flex", alignItems: "center", gap: 10, background: PANEL, border: `1px solid ${LINE}`, borderLeft: `4px solid ${ev.color}`, borderRadius: 8, padding: "8px 12px" }}>
                      <div style={{ fontSize: 12, color: "#8B93A1", width: 90, flexShrink: 0 }}>{ev.startTime}–{ev.endTime}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13.5, fontWeight: 600 }}>{ev.title} <span style={{ color: "#5C6470", fontWeight: 400 }}>· {ev.type}</span></div>
                        {ev.location && <div style={{ fontSize: 11.5, color: "#5C6470" }}>{ev.location}</div>}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}

          {pd.stats && (
            <>
              <HomeSectionLink eyebrow="Box score" title="Your totals" onClick={() => goToPlayer("stats")} />
              {box.loading ? <EmptyState text="Loading…" /> : box.entries.length === 0 ? (
                <EmptyState text="No box score imported yet." />
              ) : (
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
                  {box.statLabels.slice(0, 6).map(l => (
                    <StatPill key={l} label={friendlyStatLabel(l)} value={formatStatValue(l, box.averages[l])} />
                  ))}
                </div>
              )}

              <HomeSectionLink eyebrow="Coding file" title="Playtypes & shooting selection" onClick={() => goToPlayer("stats")} />
              {off.length === 0 && def.length === 0 ? (
                <EmptyState text="No action coded yet (Import Match tab)." />
              ) : (
                <OffenseDefenseBreakdown off={off} def={def} detailTables={false} />
              )}
            </>
          )}

          {pd.training && (
            <>
              <HomeSectionLink eyebrow="Training" title="Last 3 sessions" onClick={() => goToPlayer("training")} />
              {trainings.length === 0 ? <EmptyState text="No session recorded." /> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 26 }}>
                  {trainings.map(t => (
                    <div key={t.id} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 13 }}>{t.date} <span style={{ color: "#5C6470" }}>· {t.thematique}</span> — {t.theme || t.objectif}</span>
                      <b style={{ color: ratingColor(t.eval) }}>{t.eval}/5</b>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}

          {pd.objectives && (
            <>
              <HomeSectionLink eyebrow="Goals" title="Objectives" onClick={() => goToPlayer("objectifs")} />
              {objectives.loading ? <EmptyState text="Loading…" /> : objectives.objectives.length === 0 ? (
                <EmptyState text="No objective defined yet." />
              ) : (
                <div style={{ fontSize: 13, color: "#8B93A1", marginBottom: 26 }}>{objectives.objectives.length} active objective{objectives.objectives.length !== 1 ? "s" : ""} — tap to see progress.</div>
              )}
            </>
          )}

          {pd.mental && (
            <>
              <HomeSectionLink eyebrow="Self-assessment" title="Mental evaluation" onClick={() => goToPlayer("mental")} />
              <div style={{ fontSize: 13, color: "#8B93A1", marginBottom: 26 }}>
                {mentalEntries.length === 0 ? "No evaluation recorded yet." : `Average score: ${mentalAvg.toFixed(1)}/5 over ${mentalEntries.length} evaluation${mentalEntries.length !== 1 ? "s" : ""}.`}
              </div>
            </>
          )}
        </>
        );
      })()}

      {(isCoach || (visibility || DEFAULT_VISIBILITY).tabs.team) && (
        <>
          {(isCoach || (visibility || DEFAULT_VISIBILITY).team.advanced) && (
            <>
              <HomeSectionLink eyebrow="Team" title="Four Factors" onClick={goToTeam} />
              {advanced.loading ? <EmptyState text="Loading…" /> : !advanced.perMatch.length ? (
                <EmptyState text="No box score found in memory. Import a file from the 'Full Stats' tab (top menu)." />
              ) : (
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
                  <StatPill label="ORTG" value={ortg !== null ? ortg.toFixed(1) : "–"} tone="teal" />
                  <StatPill label="DRTG" value={drtg !== null ? drtg.toFixed(1) : "–"} tone="red" />
                  <StatPill label="eFG%" value={efg !== null ? efg.toFixed(1) + "%" : "–"} />
                  <StatPill label="TOV%" value={tovPct !== null ? tovPct.toFixed(1) + "%" : "–"} tone="red" />
                  <StatPill label="FTA/FGA" value={ftRate !== null ? ftRate.toFixed(2) : "–"} />
                  <StatPill label="OREB%" value={orebPct !== null ? orebPct.toFixed(1) + "%" : "–"} />
                </div>
              )}
            </>
          )}

          {(isCoach || (visibility || DEFAULT_VISIBILITY).team.teamPlay) && (
            <>
              <HomeSectionLink eyebrow="Coding file" title="Team playtypes & shooting selection" onClick={goToTeam} />
              {teamOff.length === 0 && teamDef.length === 0 ? (
                <EmptyState text="No action coded yet (Import Match tab)." />
              ) : (
                <OffenseDefenseBreakdown off={teamOff} def={teamDef} detailTables={false} />
              )}
            </>
          )}
        </>
      )}
    </div>
  );
}

function BottomTabBar({ tab, setTab, isCoach, visibility }) {
  const v = visibility || DEFAULT_VISIBILITY;
  const items = [
    { id: "home", label: "Home", icon: Home },
    ...(isCoach || v.tabs.planning ? [{ id: "planning", label: "Planning", icon: Calendar }] : []),
    ...(isCoach ? [{ id: "import", label: "Import", icon: Upload }] : []),
    ...(isCoach ? [{ id: "boxscore", label: "Full Stats", icon: ClipboardList }] : []),
    ...(isCoach || v.tabs.players ? [{ id: "players", label: "Players", icon: Users }] : []),
    ...(isCoach || v.tabs.team ? [{ id: "team", label: "Team", icon: LayoutGrid }] : []),
    ...(isCoach || v.tabs.scouting ? [{ id: "scouting", label: "Scouting", icon: Search }] : []),
    ...(isCoach ? [{ id: "backup", label: "Backup", icon: Download }] : []),
    ...(isCoach ? [{ id: "settings", label: "Settings", icon: ClipboardList }] : []),
  ];
  return (
    <div style={{ position: "fixed", left: 0, right: 0, bottom: 0, zIndex: 100 }}>
      <div style={{
        background: "#0B0D11", borderTop: `1px solid ${LINE}`,
        display: "flex", overflowX: "auto", WebkitOverflowScrolling: "touch", paddingBottom: "env(safe-area-inset-bottom, 0px)",
      }}>
        {items.map(it => (
          <button key={it.id} onClick={() => setTab(it.id)} style={{
            flex: "0 0 auto", width: 72, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
            gap: 3, padding: "9px 4px 8px", background: "none", border: "none", cursor: "pointer", fontFamily: "inherit",
            color: tab === it.id ? AMBER : "#7B8390",
          }}>
            <it.icon size={20} strokeWidth={tab === it.id ? 2.4 : 1.9} />
            <span style={{ fontSize: 10.5, fontWeight: tab === it.id ? 700 : 500, whiteSpace: "nowrap" }}>{it.label}</span>
          </button>
        ))}
      </div>
      {/* Léger indice visuel qu'il y a d'autres onglets à faire glisser — sans bloquer le tap. */}
      <div style={{ position: "absolute", top: 0, right: 0, bottom: "env(safe-area-inset-bottom, 0px)", width: 28, background: "linear-gradient(90deg, transparent, #0B0D11)", pointerEvents: "none" }} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

function ImportTab({ roster, onImported, matchesIndex, onDeleteMatch }) {
  const [preview, setPreview] = useState(null);
  const [fileErr, setFileErr] = useState("");
  const [date, setDate] = useState(todayLocal());
  const [opponent, setOpponent] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [requestedIds, setRequestedIds] = useState(new Set());
  const fileRef = useRef();

  async function handleDelete(m) {
    await onDeleteMatch(m.id, `${m.date} vs ${m.opponent}`);
    setRequestedIds(s => new Set([...s, m.id]));
    setConfirmDeleteId(null);
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileErr(""); setPreview(null);
    try {
      await syncPlayerCategoryFromRoster(roster);
      const buf = await file.arrayBuffer();
      const parsed = parseMatchFile(buf);
      if (!parsed.plays.length) throw new Error("No action attributed to a roster player was found.");
      // Garde une copie du fichier original en base64, pour pouvoir le rouvrir plus tard —
      // taille raisonnable pour un fichier de coding (quelques dizaines de Ko en général).
      const fileDataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      setPreview({ ...parsed, fileName: file.name, fileDataUrl });
    } catch (err) {
      setFileErr(err.message || "Error reading the file.");
    }
  }

  async function confirmImport() {
    if (!opponent.trim()) { setFileErr("Enter the opponent before confirming."); return; }
    setBusy(true);
    await onImported(preview, { date, opponent: opponent.trim() });
    setBusy(false);
    setPreview(null); setOpponent("");
    if (fileRef.current) fileRef.current.value = "";
  }

  async function openFile(id) {
    const f = await storeGet("match_file:" + id);
    if (!f) return;
    const a = document.createElement("a");
    a.href = f.dataUrl;
    a.download = f.name;
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  return (
    <div>
      <SectionTitle eyebrow="01 — Raw data" title="Import a match" />
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 22, marginBottom: 24 }}>
        <p style={{ color: "#8B93A1", fontSize: 13.5, lineHeight: 1.6, margin: "0 0 16px" }}>
          Drop the Excel file exported by the coding software (<b>Database</b> sheet).
          The useful part is detected automatically: everything before the first empty column
          after the "button" column.
        </p>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
          <div>
            <label style={labelStyle}>Match date</label>
            <input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 180 }} />
          </div>
          <div style={{ flex: 1, minWidth: 200 }}>
            <label style={labelStyle}>Opponent</label>
            <input type="text" placeholder="e.g. Zalgiris" value={opponent} onChange={e => setOpponent(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ color: "#8B93A1", fontSize: 13 }} />
        {fileErr && <div style={{ color: RED, fontSize: 13, marginTop: 10 }}>{fileErr}</div>}

        {preview && (
          <div style={{ marginTop: 18, padding: 16, background: PANEL2, borderRadius: 10, border: `1px solid ${LINE}` }}>
            <div style={{ fontSize: 13, color: PAPER, marginBottom: 10 }}>
              Sheet read: <b>{preview.sheetName}</b> · {preview.columnsDetected} columns detected before separator ·
              {" "}<b>{preview.playsWithPlayer}</b> actions attributed to a player / {preview.totalRows} total rows
            </div>
            <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 12 }}>
              <b>{preview.detectedPlayers.length}</b> players detected in the file: {preview.detectedPlayers.join(", ")}
            </div>
            {preview.unconfirmedPlayers.length > 0 && (
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 10, background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 8, marginBottom: 12 }}>
                <AlertTriangle size={14} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 12, color: "#D8DCE2", lineHeight: 1.5 }}>
                  <b>{preview.unconfirmedPlayers.join(", ")}</b> {preview.unconfirmedPlayers.length > 1 ? "are" : "is"} not in the "Player"
                  list (Settings tab) — treated as a player by default because {preview.unconfirmedPlayers.length > 1 ? "they are" : "it is"} not
                  recognized as a known tag either. If this is indeed a player, add them to the "Player" category in Settings so
                  they're recognized with certainty next time.
                </div>
              </div>
            )}
            <button disabled={busy} onClick={confirmImport} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>
              {busy ? "Import…" : "Confirm import"}
            </button>
          </div>
        )}
      </div>

      <SectionTitle eyebrow="02 — History" title="Imported matches" />
      {matchesIndex.length === 0 ? (
        <EmptyState text="No match imported yet." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {matchesIndex.slice().reverse().map(m => (
            <div key={m.id} style={{ ...btnRow, cursor: "default" }}>
              <span>{m.date} <span style={{ color: "#5C6470" }}>vs</span> {m.opponent} <span style={{ color: "#5C6470" }}>· {m.playsCount} actions</span></span>
              {requestedIds.has(m.id) ? (
                <span style={{ fontSize: 11.5, color: AMBER }}>Pending admin approval</span>
              ) : confirmDeleteId === m.id ? (
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: 11.5, color: RED }}>Delete this match?</span>
                  <button onClick={() => handleDelete(m)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                  <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                </div>
              ) : (
                <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                  <button onClick={() => openFile(m.id)} title="Open the original file" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Download size={15} /></button>
                  <button onClick={() => setConfirmDeleteId(m.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={15} /></button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const labelStyle = { display: "block", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#5C6470", marginBottom: 6 };

function SectionTitle({ eyebrow, title }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 11, color: AMBER, letterSpacing: "0.1em", marginBottom: 4 }}>{eyebrow}</div>
      <h2 style={{ fontSize: 20, fontWeight: 800, margin: 0, letterSpacing: "-0.01em" }}>{title}</h2>
    </div>
  );
}
function EmptyState({ text }) {
  return <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5 }}>{text}</div>;
}

// ---------------------------------------------------------------------------
// Joueurs — liste
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Comparaison directe de deux joueurs (n'importe lesquels, même poste ou non)
// ---------------------------------------------------------------------------

function lowerIsBetterGuess(label) {
  return /perte|faute|tov|turnover|encaiss/i.test(label);
}

function computeAveragesFromEntries(entries) {
  const statLabels = filterRedundantRawPctColumns(prioritizeLabels(Array.from(new Set(entries.flatMap(e => Object.keys(e.stats))))));
  const weighted = computeWeightedPlayerPercentages(entries);
  const WEIGHTED_LABELS = {
    "% 2pts": "pct2", "% 2pts (calculated)": "pct2", "% 3pts": "pct3", "% 3pts (calculated)": "pct3",
    "% LF": "pctFT", "% LF (calculated)": "pctFT", "eFG% (calculated)": "efg", "Usage%": "usagePct",
  };
  const averages = {};
  statLabels.forEach(l => {
    const vals = entries.map(e => e.stats[l]).filter(v => v !== undefined);
    const naiveAvg = vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
    // BUG RÉEL CORRIGÉ : si le fichier n'a pas de colonnes détaillées (tirs réussis/manqués
    // séparés) pour recalculer ce pourcentage à partir des totaux, le calcul pondéré renvoie
    // null — dans ce cas on retombe sur la moyenne simple plutôt que d'afficher "–" alors
    // qu'une vraie valeur existait pour ce match (constaté : un joueur à 1 seul match et un
    // vrai pourcentage de tir s'affichait vide).
    if (WEIGHTED_LABELS[l]) { averages[l] = weighted[WEIGHTED_LABELS[l]] ?? naiveAvg; return; }
    averages[l] = naiveAvg;
  });
  return { statLabels, averages, games: entries.length };
}

function PlayerVsPlayer({ roster, onClose }) {
  // Volontairement indépendant du filtre de saison de la page : chaque joueur a son propre
  // sélecteur de saison ci-dessous, pour pouvoir comparer deux saisons différentes.
  const fullBox = useAllBoxScores(null);
  const [playerA, setPlayerA] = useState("");
  const [playerB, setPlayerB] = useState("");
  const [seasonA, setSeasonA] = useState("all");
  const [seasonB, setSeasonB] = useState("all");
  const [extraA, setExtraA] = useState({ training: null, mental: null });
  const [extraB, setExtraB] = useState({ training: null, mental: null });
  const names = roster.filter(p => fullBox.byPlayer[p.name]);

  // Note moyenne d'entraînement et d'évaluation mentale — en plus des stats de box score.
  async function loadExtras(playerName) {
    const training = (await storeGet("training:" + playerName)) || [];
    const rated = training.filter(e => e.eval !== null && e.eval !== undefined);
    const trainingAvg = rated.length ? rated.reduce((s, e) => s + Number(e.eval), 0) / rated.length : null;
    const mental = (await storeGet("mental:" + playerName)) || [];
    const mentalAvg = mental.length
      ? mental.reduce((s, e) => s + (Object.values(e.ratings || {}).reduce((a, b) => a + b, 0) / (Object.values(e.ratings || {}).length || 1)), 0) / mental.length
      : null;
    return { training: trainingAvg, mental: mentalAvg };
  }
  useEffect(() => { if (playerA) loadExtras(playerA).then(setExtraA); else setExtraA({ training: null, mental: null }); }, [playerA]);
  useEffect(() => { if (playerB) loadExtras(playerB).then(setExtraB); else setExtraB({ training: null, mental: null }); }, [playerB]);

  function seasonsFor(first) {
    const entries = fullBox.byPlayer[first]?.entries || [];
    return Array.from(new Set(entries.map(e => e.season).filter(Boolean)));
  }
  function statsFor(first, season) {
    const entries = fullBox.byPlayer[first]?.entries || [];
    const filtered = season === "all" ? entries : entries.filter(e => e.season === season);
    return computeAveragesFromEntries(filtered);
  }

  const a = playerA ? statsFor(playerA, seasonA) : null;
  const b = playerB ? statsFor(playerB, seasonB) : null;
  const labels = Array.from(new Set([...(a?.statLabels || []), ...(b?.statLabels || [])]));
  const seasonsA = playerA ? seasonsFor(playerA) : [];
  const seasonsB = playerB ? seasonsFor(playerB) : [];

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 14, fontWeight: 700 }}>Compare two players</div>
        <button onClick={onClose} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><X size={16} /></button>
      </div>
      <div style={{ display: "flex", gap: 12, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 180 }}>
          <label style={labelStyle}>Player A</label>
          <select value={playerA} onChange={e => { setPlayerA(e.target.value); setSeasonA("all"); }} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
            <option value="">— Choose —</option>
            {names.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
          </select>
          {seasonsA.length > 0 && (
            <select value={seasonA} onChange={e => setSeasonA(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", marginTop: 6, fontSize: 12.5 }}>
              <option value="all">All seasons</option>
              {seasonsA.sort().reverse().map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          )}
        </div>
        <div style={{ flex: 1, minWidth: 180 }}>
          <label style={labelStyle}>Player B</label>
          <select value={playerB} onChange={e => { setPlayerB(e.target.value); setSeasonB("all"); }} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
            <option value="">— Choose —</option>
            {names.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
          </select>
          {seasonsB.length > 0 && (
            <select value={seasonB} onChange={e => setSeasonB(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", marginTop: 6, fontSize: 12.5 }}>
              <option value="all">All seasons</option>
              {seasonsB.sort().reverse().map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          )}
        </div>
      </div>

      {!playerA || !playerB ? (
        <EmptyState text="Choose two players to compare (any players, even different positions, even different seasons)." />
      ) : labels.length === 0 && extraA.training === null && extraA.mental === null && extraB.training === null && extraB.mental === null ? (
        <EmptyState text="No box score, training, or mental evaluation for one of these two players in the selected season." />
      ) : (
        <div style={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, overflow: "hidden" }}>
          <div style={{ display: "grid", gridTemplateColumns: "0.6fr 1fr 0.6fr", padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
            <div style={{ textAlign: "right", color: AMBER }}>{roster.find(p => p.name === playerA)?.name} <span style={{ color: "#5C6470" }}>({seasonA === "all" ? "all seasons" : seasonA}, {a.games} games)</span></div>
            <div style={{ textAlign: "center" }}>Stat</div>
            <div style={{ color: TEAL }}>{roster.find(p => p.name === playerB)?.name} <span style={{ color: "#5C6470" }}>({seasonB === "all" ? "all seasons" : seasonB}, {b.games} games)</span></div>
          </div>
          {[["Training rating (avg.)", extraA.training, extraB.training, v => v !== null ? v.toFixed(1) + "/5" : "–"],
            ["Mental evaluation (avg.)", extraA.mental, extraB.mental, v => v !== null ? v.toFixed(1) + "/5" : "–"]].map(([label, va, vb, fmt]) => {
            if (va === null && vb === null) return null;
            const aWins = va !== null && vb !== null && va !== vb && va > vb;
            const bWins = va !== null && vb !== null && va !== vb && vb > va;
            return (
              <div key={label} style={{ display: "grid", gridTemplateColumns: "0.6fr 1fr 0.6fr", padding: "8px 16px", alignItems: "center", borderBottom: `1px solid ${LINE}`, fontSize: 13, background: PANEL }}>
                <div style={{ textAlign: "right", fontFamily: "ui-monospace, monospace", fontWeight: 700, color: aWins ? TEAL : PAPER }}>{fmt(va)}</div>
                <div style={{ textAlign: "center", color: "#8B93A1", fontSize: 12 }}>{label}</div>
                <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, color: bWins ? TEAL : PAPER }}>{fmt(vb)}</div>
              </div>
            );
          })}
          {labels.map(l => {
            const va = a?.averages[l], vb = b?.averages[l];
            if (va === undefined && vb === undefined) return null;
            const lowerBetter = lowerIsBetterGuess(l);
            const aWins = va !== undefined && vb !== undefined && va !== vb && (lowerBetter ? va < vb : va > vb);
            const bWins = va !== undefined && vb !== undefined && va !== vb && (lowerBetter ? vb < va : vb > va);
            return (
              <div key={l} style={{ display: "grid", gridTemplateColumns: "0.6fr 1fr 0.6fr", padding: "8px 16px", alignItems: "center", borderBottom: `1px solid ${LINE}`, fontSize: 13 }}>
                <div style={{ textAlign: "right", fontFamily: "ui-monospace, monospace", fontWeight: 700, color: aWins ? TEAL : PAPER }}>{formatStatValue(l, va)}</div>
                <div style={{ textAlign: "center", color: "#8B93A1", fontSize: 12 }}>{l}</div>
                <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, color: bWins ? TEAL : PAPER }}>{formatStatValue(l, vb)}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function PlayersList({ roster, allPlays, onSelect, isCoach, onlyOwn, matchFilter, onAddPlayer, onRemovePlayer, onEditPlayer }) {
  const visibleRoster = onlyOwn ? roster.filter(p => p.name === onlyOwn) : roster;
  const box = useAllBoxScores(matchFilter);
  const [showAdd, setShowAdd] = useState(false);
  const [showCompare, setShowCompare] = useState(false);
  const [newName, setNewName] = useState("");
  const [newPosition, setNewPosition] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(null); // id du joueur en confirmation de suppression
  const [editingId, setEditingId] = useState(null);
  const [editName, setEditName] = useState("");
  const [editPosition, setEditPosition] = useState("");
  const [busy, setBusy] = useState(false);
  const [requestedDeletionIds, setRequestedDeletionIds] = useState(new Set());

  async function handleDeletePlayer(p) {
    await onRemovePlayer(p.id, p.name);
    setRequestedDeletionIds(s => new Set([...s, p.id]));
    setConfirmDelete(null);
  }

  const rows = visibleRoster.map(p => {
    const plays = allPlays.filter(pl => pl.player === p.name);
    const codedGames = new Set(plays.map(pl => pl.matchId)).size;
    const b = box.byPlayer[p.name];
    const ppg = b && b.ptsLabel && b.averages[b.ptsLabel] !== null ? b.averages[b.ptsLabel].toFixed(1) : "–";
    return { ...p, playCount: plays.length, codedGames, boxGames: b ? b.games : 0, ppg };
  });

  async function submitAdd() {
    if (!newName.trim()) return;
    setBusy(true);
    await onAddPlayer(newName, newPosition);
    setBusy(false); setNewName(""); setNewPosition(""); setShowAdd(false);
  }

  function startEdit(p) { setEditingId(p.id); setEditName(p.name); setEditPosition(p.position || ""); }

  async function submitEdit(p) {
    if (!editName.trim()) return;
    const newFirst = editName.trim().split(" ")[0];
    setBusy(true);
    await onEditPlayer(p.id, { name: editName.trim(), first: newFirst, position: editPosition });
    setBusy(false); setEditingId(null);
  }

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <SectionTitle eyebrow="Roster" title={onlyOwn ? "My profile" : "Players"} />
        <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
          {!onlyOwn && (
            <button onClick={() => setShowCompare(o => !o)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: TEAL, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
              <Users size={14} /> Compare 2 players
            </button>
          )}
          {isCoach && !onlyOwn && (
            <button onClick={() => setShowAdd(o => !o)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: AMBER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
              <Plus size={14} /> Add a player
            </button>
          )}
        </div>
      </div>

      {showCompare && <PlayerVsPlayer roster={roster} onClose={() => setShowCompare(false)} />}

      {showAdd && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div style={{ flex: 1, minWidth: 180 }}>
            <label style={labelStyle}>Full name</label>
            <input value={newName} onChange={e => setNewName(e.target.value)} placeholder="e.g. Jean Dupont" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
          <div style={{ width: 180 }}>
            <label style={labelStyle}>Position</label>
            <select value={newPosition} onChange={e => setNewPosition(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
              <option value="">—</option>
              <option value="Point Guard">Point Guard</option>
              <option value="Forward">Forward</option>
              <option value="Big">Big</option>
            </select>
          </div>
          <button disabled={busy} onClick={submitAdd} style={{ ...btnPrimary, width: "auto", padding: "11px 18px" }}>{busy ? "…" : "Add"}</button>
          <button onClick={() => setShowAdd(false)} style={btnSecondary}>Cancel</button>
        </div>
      )}

      {rows.length === 0 && <div style={{ marginBottom: 16 }}><EmptyState text="No players in the roster yet — add the first player above." /></div>}
      {rows.length > 0 && rows.every(r => r.playCount === 0 && r.boxGames === 0) && <div style={{ marginBottom: 16 }}><EmptyState text="No match data imported yet — stats will appear here after an import." /></div>}
      <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 6 }}>
        {rows.map(p => (
          editingId === p.id ? (
            <div key={p.id} style={{ background: PANEL2, border: `1px solid ${AMBER}`, borderRadius: 10, padding: 16 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 10 }}>
                <div style={{ flex: 1, minWidth: 180 }}>
                  <label style={labelStyle}>Full name</label>
                  <input value={editName} onChange={e => setEditName(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
                </div>
                <div style={{ width: 180 }}>
                  <label style={labelStyle}>Position</label>
                  <select value={editPosition} onChange={e => setEditPosition(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
                    <option value="">—</option>
                    <option value="Point Guard">Point Guard</option>
                    <option value="Forward">Forward</option>
                    <option value="Big">Big</option>
                  </select>
                </div>
              </div>
              {editName.trim() !== p.name && (
                <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 10, background: PANEL, border: `1px solid ${TEAL}`, borderRadius: 8, marginBottom: 10 }}>
                  <ShieldCheck size={14} color={TEAL} style={{ marginTop: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: 11.5, color: "#D8DCE2", lineHeight: 1.4 }}>
                    The name is changing ("{p.name}" → "{editName.trim()}"). Everything already recorded for this player
                    (training, objectives, mental evaluations, Wellness, role, meetings, coded matches, box scores, login)
                    will automatically move over to the new name — nothing gets lost.
                  </div>
                </div>
              )}
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => setEditingId(null)} style={btnSecondary}>Cancel</button>
                <button disabled={busy} onClick={() => submitEdit(p)} style={{ ...btnPrimary, width: "auto", padding: "9px 16px" }}>{busy ? "…" : "Save"}</button>
              </div>
            </div>
          ) : (
            <div key={p.id} style={{ ...btnRow, cursor: "default" }}>
              <button onClick={() => onSelect(p.name)} style={{ display: "flex", alignItems: "center", gap: 14, background: "none", border: "none", cursor: "pointer", padding: 0, flex: 1, textAlign: "left", fontFamily: "inherit", color: "inherit" }}>
                <PlayerAvatar playerName={p.first} size={34} />
                <div>
                  <div style={{ fontWeight: 600, fontSize: 14.5 }}>{p.name}</div>
                  <div style={{ fontSize: 12, color: "#5C6470" }}>{p.position ? p.position + " · " : ""}{p.boxGames} box score{p.boxGames !== 1 ? "s" : ""} · {p.codedGames} match{p.codedGames !== 1 ? "s" : ""} coded</div>
                </div>
              </button>
              <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                <button onClick={() => onSelect(p.name)} style={{ background: "none", border: "none", cursor: "pointer", display: "flex", padding: 0 }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 16, fontWeight: 700, color: AMBER }}>{p.ppg}</div>
                    <div style={{ fontSize: 10, color: "#5C6470", textTransform: "uppercase" }}>pts/game (box score)</div>
                  </div>
                </button>
                {isCoach && !onlyOwn && (
                  requestedDeletionIds.has(p.id) ? (
                    <span style={{ fontSize: 11.5, color: AMBER }}>Pending admin approval</span>
                  ) : confirmDelete === p.id ? (
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <span style={{ fontSize: 11.5, color: RED }}>Delete?</span>
                      <button onClick={(e) => { e.stopPropagation(); handleDeletePlayer(p); }} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11.5, padding: "5px 9px", cursor: "pointer" }}>Yes</button>
                      <button onClick={(e) => { e.stopPropagation(); setConfirmDelete(null); }} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11.5, padding: "5px 9px", cursor: "pointer" }}>No</button>
                    </div>
                  ) : (
                    <>
                      <button onClick={(e) => { e.stopPropagation(); startEdit(p); }} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }} title="Edit"><ClipboardList size={15} /></button>
                      <button onClick={(e) => { e.stopPropagation(); setConfirmDelete(p.id); }} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={15} /></button>
                    </>
                  )
                )}
                <ChevronRight size={16} color="#5C6470" onClick={() => onSelect(p.name)} style={{ cursor: "pointer" }} />
              </div>
            </div>
          )
        ))}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Joueur — détail
// ---------------------------------------------------------------------------

function PositionComparisonTable({ position, roster, byPlayer, playerFirst }) {
  const peers = roster.filter(p => p.position === position);
  if (peers.length < 2) return <EmptyState text="No other player recorded at this position yet." />;

  const keys = POSITION_FEATURED_STATS[position] || ["pts", "tov"];
  // Union des colonnes réellement présentes chez au moins un des joueurs du poste, dans
  // l'ordre de priorité défini pour ce poste (Pts, Assists, Ballons perdus, % LF, % 3Pts…).
  const allLabelsAtPosition = Array.from(new Set(peers.flatMap(p => (byPlayer[p.name]?.statLabels) || [])));
  // BUG RÉEL CORRIGÉ : pour les pourcentages de tir (twoPct/tpmPct/ftPct), utiliser la colonne
  // BRUTE du fichier ("3PTS%", "3PT %"…) plutôt que la version normalisée ("% 3pts") pose
  // problème dès que deux fichiers importés n'utilisent pas exactement le même intitulé de
  // colonne — un seul des deux matchs était alors pris en compte dans la moyenne (constaté :
  // Nathan Mettler affiché à 50% au lieu de 25% sur ses deux matchs). "% 2pts"/"% 3pts"/"% LF"
  // sont TOUJOURS le même nom quel que soit le fichier d'origine, donc on les priorise ici.
  const NORMALIZED_PCT_LABEL = { twoPct: "% 2pts", tpmPct: "% 3pts", ftPct: "% LF" };
  const columns = keys.map(key => {
    if (NORMALIZED_PCT_LABEL[key] && allLabelsAtPosition.includes(NORMALIZED_PCT_LABEL[key])) {
      return { key, label: NORMALIZED_PCT_LABEL[key], fallbackLabel: STAT_KEY_LABEL_FR[key] || key };
    }
    return { key, label: findStatCol(allLabelsAtPosition, STAT_PATTERNS[key] || [], key), fallbackLabel: STAT_KEY_LABEL_FR[key] || key };
  });

  const rows = peers.map(p => {
    const b = byPlayer[p.name];
    return { name: p.name, first: p.first, values: columns.map(c => (c.label && b && b.averages[c.label] !== undefined) ? b.averages[c.label] : null) };
  });

  const primaryIdx = 0;
  rows.sort((a, b) => (b.values[primaryIdx] ?? -Infinity) - (a.values[primaryIdx] ?? -Infinity));

  if (!columns.some(c => c.label)) return <EmptyState text="No box score yet for players at this position." />;

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
      <div style={{ display: "grid", gridTemplateColumns: `1.4fr ${columns.map(() => "0.8fr").join(" ")}`, padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
        <div>Player ({position})</div>
        {columns.map(c => <div key={c.key}>{c.label || c.fallbackLabel}</div>)}
      </div>
      {rows.map(r => (
        <div key={r.first} style={{
          display: "grid", gridTemplateColumns: `1.4fr ${columns.map(() => "0.8fr").join(" ")}`, padding: "10px 16px", alignItems: "center",
          borderBottom: `1px solid ${LINE}`, fontSize: 13.5, background: r.name === playerFirst ? PANEL2 : "transparent",
        }}>
          <div style={{ fontWeight: r.name === playerFirst ? 700 : 500, color: r.name === playerFirst ? AMBER : PAPER }}>{r.name}</div>
          {r.values.map((v, i) => (
            <div key={i} style={{ fontFamily: "ui-monospace, monospace", color: "#8B93A1" }}>{v !== null ? formatStatValue(columns[i].label || columns[i].key, v) : "–"}</div>
          ))}
        </div>
      ))}
    </div>
  );
}

// Rapport imprimable (utilisé par le bouton "Export PDF" via l'impression navigateur —
// aucune librairie PDF n'est disponible dans cet environnement, donc on passe par
// window.print() avec une mise en page dédiée).

function PlayerPrintReport({ playerName, position, off, def, box, allBox, roster }) {
  const [wellnessEntries, setWellnessEntries] = useState([]);
  useEffect(() => { storeGet("wellness:" + playerName).then(w => setWellnessEntries(w || [])); }, [playerName]);
  const wellnessChartData = [...wellnessEntries].sort((a, b) => wellnessSortKey(a).localeCompare(wellnessSortKey(b)))
    .map(e => ({ match: e.date + " " + (WELLNESS_SLOTS.find(s => s.key === e.slot)?.label || e.slot).slice(0, 3), physical: e.physical, mental: e.mental }));

  return (
    <div style={{ padding: 24, background: "#ffffff", color: "#1A1D24" }}>
      <h1 style={{ fontSize: 24, marginBottom: 4 }}>{playerName}</h1>
      <div style={{ fontSize: 12, color: "#8B93A1", marginBottom: 20 }}>{position || "Position not set"} · Report generated on {new Date().toLocaleDateString("en-US")}</div>

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Official totals (box score)</h2>
      {box.entries.length === 0 ? <p>No box score imported.</p> : (
        <>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
            <StatPill label="Games played" value={box.entries.length} sub="games he actually played in (box score)" />
            {(() => {
              // Même logique EXACTEMENT que l'écran (même plafond de 8) — l'export doit refléter
              // ce que montre le site, pas afficher en plus des statistiques brutes qui n'y
              // apparaissent jamais (constaté : 31 encadrés à l'export contre 9 à l'écran).
              const featured = position ? featuredStatsForPosition(position, box.statLabels).filter(f => f.label) : [];
              const featuredLabels = new Set(featured.map(f => f.label));
              const rest = box.statLabels.filter(l => !featuredLabels.has(l)).slice(0, 8 - featured.length);
              return (
                <>
                  {featured.map(f => {
                    const rank = teamRank(allBox.byPlayer, f.label, playerName);
                    return <StatPill key={f.key} label={f.label} value={formatStatValue(f.label, box.averages[f.label])} sub={rank ? `${position} · #${rank.rank} of ${rank.total} team` : position} tone="amber" />;
                  })}
                  {rest.map(l => {
                    const rank = teamRank(allBox.byPlayer, l, playerName);
                    return <StatPill key={l} label={l} value={formatStatValue(l, box.averages[l])} sub={rank ? `average / game · #${rank.rank} of ${rank.total} team` : "average / game"} tone="teal" />;
                  })}
                </>
              );
            })()}
          </div>

          {(() => {
            // Manquait à l'export alors qu'affiché à l'écran — mêmes données, même source
            // (box.averages, déjà calculé correctement via derivedMatchStats/useBoxScore).
            const minutesLabel = findStatCol(box.statLabels, STAT_PATTERNS.minutes, "minutes");
            const playerMinutes = minutesLabel ? box.averages[minutesLabel] : undefined;
            const theoreticalPoss = box.averages["Theoretical possessions"];
            const usagePct = box.averages["Usage%"];
            return (
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
                <StatPill label="Playing time" value={playerMinutes !== undefined ? playerMinutes.toFixed(1) + " min" : "–"} sub={minutesLabel ? "average / game" : "data missing from box score"} />
                <StatPill label="Possessions played (theoretical)" value={theoreticalPoss !== undefined ? theoreticalPoss.toFixed(1) : "–"} sub={playerMinutes !== undefined ? "based on team possessions & playing time" : "requires playing time"} />
                <StatPill label="% Usage" value={usagePct !== undefined && usagePct !== null ? usagePct.toFixed(1) + "%" : "–"} sub={playerMinutes !== undefined ? "possessions ended / possessions played" : "requires playing time"} tone="red" />
              </div>
            );
          })()}

          <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Match by match</h2>
          {/* Police et espacement volontairement compacts pour que le tableau tienne sur la
              largeur de la page à l'export — un tableau "overflow-x: auto" comme à l'écran ne
              fonctionne pas dans un PDF statique, il serait simplement coupé sans possibilité de défiler.
              Titres tournés à 90° (lus de bas en haut, comme un tableur) pour libérer beaucoup
              plus de largeur par colonne — les valeurs peuvent aussi revenir à la ligne si besoin
              (sinon un nombre comme "-18.0" ou "25.0%" déborde visuellement sur la colonne suivante).
              CORRECTIF : writing-mode (utilisé précédemment) n'est pas fiable avec html2canvas
              (moteur de rendu utilisé pour l'export PDF, différent d'un vrai navigateur) — le texte
              apparaissait à l'envers. Une simple rotation CSS (transform: rotate) est la technique
              standard, bien mieux supportée par ce type de moteur de rendu vers canevas. */}
          <table style={{ borderCollapse: "collapse", width: "100%", marginBottom: 12, tableLayout: "fixed" }}>
            <thead><tr style={{ height: 110 }}>
              <th style={{ border: `1px solid ${LINE}`, padding: "4px 3px", verticalAlign: "bottom" }}>
                <div style={{ transform: "rotate(-90deg)", transformOrigin: "center", fontSize: 9, fontWeight: 700, whiteSpace: "nowrap", width: 20 }}>Date</div>
              </th>
              <th style={{ border: `1px solid ${LINE}`, padding: "4px 3px", verticalAlign: "bottom" }}>
                <div style={{ transform: "rotate(-90deg)", transformOrigin: "center", fontSize: 9, fontWeight: 700, whiteSpace: "nowrap", width: 20 }}>Opponent</div>
              </th>
              {box.statLabels.map(l => (
                <th key={l} style={{ border: `1px solid ${LINE}`, padding: "4px 3px", verticalAlign: "bottom" }}>
                  <div style={{ transform: "rotate(-90deg)", transformOrigin: "center", fontSize: 9, fontWeight: 700, whiteSpace: "nowrap", width: 20 }}>{friendlyStatLabel(l)}</div>
                </th>
              ))}
            </tr></thead>
            <tbody>
              {box.entries.map((e, i) => (
                <tr key={i}>
                  <td style={{ border: `1px solid ${LINE}`, padding: "4px 3px", fontSize: 8.5, lineHeight: 1.2, whiteSpace: "normal", wordBreak: "break-word", overflow: "hidden" }}>{e.date}</td>
                  <td style={{ border: `1px solid ${LINE}`, padding: "4px 3px", fontSize: 8.5, lineHeight: 1.2, whiteSpace: "normal", wordBreak: "break-word", overflow: "hidden" }}>{e.opponent}</td>
                  {box.statLabels.map(l => <td key={l} style={{ border: `1px solid ${LINE}`, padding: "4px 3px", fontFamily: "ui-monospace, monospace", fontSize: 8.5, lineHeight: 1.2, whiteSpace: "normal", wordBreak: "break-word", overflow: "hidden" }}>{formatStatValue(l, e.stats[l])}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      {position && (
        <>
          <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Comparison by position — {position}</h2>
          <PositionComparisonTable position={position} roster={roster} byPlayer={allBox.byPlayer} playerFirst={playerName} />
        </>
      )}

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Offense / Defense</h2>
      <OffenseDefenseBreakdown off={off} def={def} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Objectives</h2>
      <ObjectivesPanel playerName={playerName} isCoach={false} box={box} off={off} def={def} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Training</h2>
      <TrainingLog playerName={playerName} isCoach={false} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Mental evaluation</h2>
      <MentalLog playerName={playerName} isCoach={false} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Wellness</h2>
      {wellnessChartData.length === 0 ? <p>No wellness entry recorded.</p> : (
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          <div style={{ width: 420, height: 200 }}>
            <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 4 }}>Physically</div>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={wellnessChartData} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
                <XAxis dataKey="match" tick={{ fill: "#5C6470", fontSize: 9 }} axisLine={{ stroke: LINE }} />
                <YAxis domain={[1, 5]} tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} />
                <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}` }} />
                <Line type="monotone" dataKey="physical" stroke={TEAL} strokeWidth={2.5} dot={{ r: 3, fill: TEAL }} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div style={{ width: 420, height: 200 }}>
            <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 4 }}>Mentally</div>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={wellnessChartData} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
                <XAxis dataKey="match" tick={{ fill: "#5C6470", fontSize: 9 }} axisLine={{ stroke: LINE }} />
                <YAxis domain={[1, 5]} tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} />
                <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}` }} />
                <Line type="monotone" dataKey="mental" stroke={AMBER} strokeWidth={2.5} dot={{ r: 3, fill: AMBER }} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Role</h2>
      <RoleTab playerName={playerName} isCoach={false} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Meetings</h2>
      <MeetingsTab playerName={playerName} isCoach={false} />
    </div>
  );
}

function PlayerDetail({ playerName, allPlays, roster, onBack, isCoach, matchFilter, visibility, teamId, teamName, initialSubtab, onEditPlayer }) {
  const v = (visibility || DEFAULT_VISIBILITY);
  const pd = v.playerDetail || DEFAULT_VISIBILITY.playerDetail;
  const plays = allPlays.filter(p => p.player === playerName);
  const off = plays.filter(isOffense);
  const def = plays.filter(isDefense);
  const games = useMemo(() => Array.from(new Set(plays.map(p => p.matchId))), [plays]);
  const box = useBoxScore(playerName, matchFilter);
  const allBox = useAllBoxScores(matchFilter);
  const teamAdvanced = useTeamAdvancedStats(matchFilter);
  const me = roster.find(p => p.name === playerName);
  const position = me ? me.position : "";
  const [editingPosition, setEditingPosition] = useState(false);
  const [positionDraft, setPositionDraft] = useState(position);

  async function savePosition() {
    if (me && onEditPlayer) await onEditPlayer(me.id, { position: positionDraft });
    setEditingPosition(false);
  }

  const SUBTAB_ORDER = ["stats", "objectifs", "training", "mental", "wellness", "role", "meetings"];
  const SUBTAB_KEY = { stats: "stats", objectifs: "objectives", training: "training", mental: "mental", wellness: "wellness", role: "role", meetings: "meetings" };
  const defaultSubtab = isCoach ? "stats" : (SUBTAB_ORDER.find(id => pd[SUBTAB_KEY[id]]) || "stats");
  const [subtab, setSubtab] = useState(initialSubtab || defaultSubtab);
  const [exportReport, setExportReport] = useState(null);
  const [trendStats, setTrendStats] = useState([]);
  const [trendMenuOpen, setTrendMenuOpen] = useState(false);

  return (
    <div>
      {isCoach && (
        <button onClick={onBack} className="screen-only" style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: "#8B93A1", cursor: "pointer", fontSize: 13, padding: 0, marginBottom: 18 }}>
          <ChevronLeft size={15} /> Back to players
        </button>
      )}

      <div className="screen-only" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16, marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <PlayerAvatar playerName={playerName} size={72} editable={isCoach} />
          <div>
            <h2 style={{ fontSize: 26, fontWeight: 800, margin: 0, letterSpacing: "-0.01em" }}>{playerName}</h2>
            {isCoach ? (
              editingPosition ? (
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4 }}>
                  <select value={positionDraft} onChange={e => setPositionDraft(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", fontSize: 12.5, padding: "5px 8px", width: "auto" }}>
                    <option value="">—</option>
                    <option value="Point Guard">Point Guard</option>
                    <option value="Forward">Forward</option>
                    <option value="Big">Big</option>
                  </select>
                  <button onClick={savePosition} style={{ fontSize: 11.5, color: TEAL, background: "none", border: "none", cursor: "pointer", fontWeight: 700 }}>Save</button>
                  <button onClick={() => { setEditingPosition(false); setPositionDraft(position); }} style={{ fontSize: 11.5, color: "#5C6470", background: "none", border: "none", cursor: "pointer" }}>Cancel</button>
                </div>
              ) : (
                <button onClick={() => { setPositionDraft(position); setEditingPosition(true); }} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: "#5C6470", background: "none", border: "none", cursor: "pointer", padding: 0, fontFamily: "inherit" }}>
                  {position || "Set position"} <ClipboardList size={11} />
                </button>
              )
            ) : (
              position && <div style={{ fontSize: 12.5, color: "#5C6470" }}>{position}</div>
            )}
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ fontSize: 12.5, color: "#5C6470" }}>{games.length} match{games.length !== 1 ? "s" : ""} coded · {box.entries.length} box score{box.entries.length !== 1 ? "s" : ""}</div>
          <button onClick={async () => {
            const filename = `fiche_${playerName}_${todayLocal()}.html`;
            const pdfOk = await tryExportPdf("player-print-content", filename, "light");
            if (pdfOk) return;
            const r = buildReportHtml("player-print-content", filename, "light");
            if (!r) { alert("Content not found — try again after the page has fully loaded."); return; }
            tryDownload(r.full, r.filename);
            setExportReport(r);
          }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
            <Download size={14} /> Export PDF
          </button>
        </div>
      </div>

      <div className="screen-only" style={{ display: "flex", gap: 8, marginBottom: 20, borderBottom: `1px solid ${LINE}`, paddingBottom: 10, flexWrap: "wrap" }}>
        {[["stats", "Match Stats", "stats"], ["objectifs", "Objectives", "objectives"], ["training", "Training", "training"], ["mental", "Mental evaluation", "mental"], ["wellness", "Wellness", "wellness"], ["role", "Role", "role"], ["meetings", "Meetings", "meetings"]]
          .filter(([, , key]) => isCoach || pd[key])
          .map(([id, label]) => (
          <button key={id} onClick={() => setSubtab(id)} style={{
            padding: "7px 12px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "inherit",
            background: subtab === id ? PANEL2 : "transparent", color: subtab === id ? AMBER : "#8B93A1"
          }}>{label}</button>
        ))}
      </div>

      <div className="screen-only">
        {!isCoach && !pd[SUBTAB_KEY[subtab]] && (
          <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5, marginBottom: 26 }}>
            You don't have the permission to see this.
          </div>
        )}
        {subtab === "stats" && (isCoach || pd.stats) && (
          <>
            <SectionTitle eyebrow="Source: Box Score" title="Official totals" />
            {box.loading ? null : box.entries.length === 0 ? (
              <div style={{ marginBottom: 26 }}><EmptyState text="No box score imported — import one from the 'Full Stats' tab to display official totals (points, %, rebounds, assists…)." /></div>
            ) : (
              <>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
                  <StatPill label="Games played" value={box.entries.length} sub="games he actually played in (box score)" />
                  {(() => {
                    const featured = position ? featuredStatsForPosition(position, box.statLabels).filter(f => f.label) : [];
                    const featuredLabels = new Set(featured.map(f => f.label));
                    const rest = box.statLabels.filter(l => !featuredLabels.has(l)).slice(0, 8 - featured.length);
                    return (
                      <>
                        {featured.map(f => {
                          const rank = teamRank(allBox.byPlayer, f.label, playerName);
                          return (
                            <StatPill key={f.key} label={f.label} value={formatStatValue(f.label, box.averages[f.label])}
                              sub={rank ? `${position} · #${rank.rank} of ${rank.total} team` : position} tone="amber" />
                          );
                        })}
                        {rest.map(l => {
                          const rank = teamRank(allBox.byPlayer, l, playerName);
                          return (
                            <StatPill key={l} label={l} value={formatStatValue(l, box.averages[l])}
                              sub={rank ? `average / game · #${rank.rank} of ${rank.total} team` : "average / game"} tone="teal" />
                          );
                        })}
                      </>
                    );
                  })()}
                </div>

                {(() => {
                  // Ces indicateurs dépendent du temps de jeu, absent de la plupart des box
                  // scores importés jusqu'ici — on affiche "–" plutôt qu'un faux calcul quand la
                  // donnée manque. Les valeurs viennent de derivedMatchStats (via useBoxScore),
                  // seule source de vérité pour ce calcul — BUG RÉEL CORRIGÉ : ce bloc avait sa
                  // propre formule dupliquée et fausse (facteur ×5 en trop), jamais mise à jour
                  // en même temps que le reste, ce qui donnait ~49.8% au lieu de ~10.8% pour le
                  // même joueur et le même match.
                  const minutesLabel = findStatCol(box.statLabels, STAT_PATTERNS.minutes, "minutes");
                  const playerMinutes = minutesLabel ? box.averages[minutesLabel] : undefined;
                  const theoreticalPoss = box.averages["Theoretical possessions"];
                  const usagePct = box.averages["Usage%"];

                  return (
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
                      <StatPill label="Playing time" value={playerMinutes !== undefined ? playerMinutes.toFixed(1) + " min" : "–"} sub={minutesLabel ? "average / game" : "data missing from box score"} />
                      <StatPill label="Possessions played (theoretical)" value={theoreticalPoss !== undefined ? theoreticalPoss.toFixed(1) : "–"} sub={playerMinutes !== undefined ? "based on team possessions & playing time" : "requires playing time"} />
                      <StatPill label="% Usage" value={usagePct !== undefined && usagePct !== null ? usagePct.toFixed(1) + "%" : "–"} sub={playerMinutes !== undefined ? "possessions ended / possessions played" : "requires playing time"} tone="red" />
                    </div>
                  );
                })()}
                <BoxScoreHistoryTable box={box} />

                {box.entries.length >= 2 && (
                  <>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
                      <SectionTitle eyebrow="Progression" title="Match-by-match evolution" />
                      <div style={{ position: "relative", marginBottom: 16 }}>
                        <button onClick={() => setTrendMenuOpen(o => !o)} style={{
                          display: "flex", alignItems: "center", gap: 8, padding: "9px 14px", background: PANEL, border: `1px solid ${LINE}`,
                          borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit",
                        }}>
                          <TrendingUp size={14} color={AMBER} />
                          {trendStats.length === 0 ? "Choose stats (max 3)" : trendStats.map(friendlyStatLabel).join(", ")}
                          <ChevronRight size={13} style={{ transform: trendMenuOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                        </button>
                        {trendMenuOpen && (
                          <div style={{ position: "absolute", right: 0, zIndex: 10, marginTop: 6, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, padding: 10, minWidth: 240, maxHeight: 320, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
                            <div style={{ fontSize: 11, color: "#5C6470", padding: "2px 8px 8px" }}>3 maximum at once</div>
                            {box.statLabels.map(l => {
                              const checked = trendStats.includes(l);
                              const disabled = !checked && trendStats.length >= 3;
                              return (
                                <button key={l} disabled={disabled} onClick={() => setTrendStats(s => checked ? s.filter(x => x !== l) : [...s, l])} style={{
                                  ...btnRow, marginBottom: 4, opacity: disabled ? 0.4 : 1, cursor: disabled ? "not-allowed" : "pointer",
                                }}>
                                  <span>{friendlyStatLabel(l)}</span>{checked && <span style={{ color: AMBER }}>✓</span>}
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 26 }}>
                      {trendStats.length === 0
                        ? <EmptyState text="Choose up to 3 stats above to see their match-by-match evolution." />
                        : trendStats.map(l => <PlayerTrendChart key={l} entries={box.entries} statLabel={l} />)}
                    </div>
                  </>
                )}
              </>
            )}

            {position && (
              <>
                <SectionTitle eyebrow={`Position : ${position}`} title="Position Comparison" />
                <div style={{ marginBottom: 26 }}>
                  <PositionComparisonTable position={position} roster={roster} byPlayer={allBox.byPlayer} playerFirst={playerName} />
                </div>
              </>
            )}

            <OffenseDefenseBreakdown off={off} def={def} />
          </>
        )}

        {subtab === "objectifs" && (isCoach || pd.objectives) && <ObjectivesPanel playerName={playerName} isCoach={isCoach} box={box} off={off} def={def} />}
        {subtab === "training" && (isCoach || pd.training) && <TrainingLog playerName={playerName} isCoach={isCoach} />}
        {subtab === "mental" && (isCoach || pd.mental) && <MentalLog playerName={playerName} isCoach={isCoach} />}
        {subtab === "wellness" && (isCoach || pd.wellness) && <WellnessTab playerName={playerName} isCoach={isCoach} canSeeCharts={isCoach || v.wellnessCharts} teamId={teamId} teamName={teamName} />}
        {subtab === "role" && (isCoach || pd.role) && <RoleTab playerName={playerName} isCoach={isCoach} />}
        {subtab === "meetings" && (isCoach || pd.meetings) && <MeetingsTab playerName={playerName} isCoach={isCoach} />}
      </div>

      <div className="print-only" id="player-print-content">
        <PlayerPrintReport playerName={playerName} position={position} off={off} def={def} box={box} allBox={allBox} roster={roster} />
      </div>
      <ExportModal report={exportReport} onClose={() => setExportReport(null)} />
    </div>
  );
}



function TagTable({ stats }) {
  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 0.6fr 0.9fr 0.6fr 0.6fr", padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
        <div>Tag</div><div>Occurrences</div><div>Frequency</div><div>PPPP</div><div>Open</div>
      </div>
      {stats.length === 0 && <div style={{ padding: 20, color: "#5C6470", fontSize: 13 }}>Not enough data yet for a breakdown.</div>}
      {stats.slice(0, 40).map(t => (
        <div key={t.tag} style={{ display: "grid", gridTemplateColumns: "1.3fr 0.6fr 0.9fr 0.6fr 0.6fr", padding: "10px 16px", alignItems: "center", borderBottom: `1px solid ${LINE}`, fontSize: 13.5 }}>
          <div>{t.tag}</div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: "#8B93A1" }}>{t.count}</div>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 70 }}><Bar pct={t.freq} /></div>
              <span style={{ fontFamily: "ui-monospace, monospace", fontSize: 12, color: "#8B93A1" }}>{t.freq.toFixed(0)}%</span>
            </div>
          </div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontWeight: 700 }}>{t.pppp.toFixed(2)}</div>
          <div style={{ fontFamily: "ui-monospace, monospace", color: t.open !== null ? TEAL : "#5C6470" }}>{t.open !== null ? t.open.toFixed(0) + "%" : "–"}</div>
        </div>
      ))}
    </div>
  );
}


// ---------------------------------------------------------------------------
// Historique box score (table détaillée match par match)
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Match-by-match evolution (courbe de tendance)
// ---------------------------------------------------------------------------

function PlayerTrendChart({ entries, statLabel }) {
  if (!statLabel) return null;
  const data = entries
    .filter(e => e.stats[statLabel] !== undefined)
    .map(e => ({ match: `${e.date}`, opponent: e.opponent, value: e.stats[statLabel] }));
  if (data.length < 2) return null;

  const avg = data.reduce((s, d) => s + d.value, 0) / data.length;
  const isPct = /%/.test(statLabel);

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "16px 10px 8px 0", flex: "1 1 320px" }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, padding: "0 18px 4px" }}>{friendlyStatLabel(statLabel)}</div>
      <div style={{ fontSize: 11, color: "#5C6470", padding: "0 18px 10px" }}>Average: {isPct ? (avg * (Math.abs(avg) <= 1 ? 100 : 1)).toFixed(1) + "%" : avg.toFixed(1)}</div>
      <ResponsiveContainer width="100%" height={180}>
        <ComposedChart data={data} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
          <CartesianGrid stroke={LINE} vertical={false} />
          <XAxis dataKey="match" tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} />
          <YAxis tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} tickFormatter={v => isPct ? formatStatValue(statLabel, v) : v} />
          <Tooltip
            contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }}
            labelFormatter={(l, p) => p?.[0]?.payload ? `${l} · vs ${p[0].payload.opponent}` : l}
            formatter={(v) => [isPct ? formatStatValue(statLabel, v) : v, statLabel]}
          />
          <Line type="monotone" dataKey="value" stroke={AMBER} strokeWidth={2.5} dot={{ r: 4, fill: AMBER }} />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}

function BoxScoreHistoryTable({ box }) {
  if (box.loading || !box.entries.length) return null;
  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflowX: "auto", marginBottom: 26 }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
        <thead>
          <tr>
            <th style={thStyle}>Date</th><th style={thStyle}>Opponent</th>
            {box.statLabels.map(l => <th key={l} style={thStyle}>{friendlyStatLabel(l)}</th>)}
          </tr>
        </thead>
        <tbody>
          {box.entries.map((e, i) => (
            <tr key={i}>
              <td style={tdStyle}>{e.date}</td><td style={tdStyle}>{e.opponent}</td>
              {box.statLabels.map(l => <td key={l} style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{formatStatValue(l, e.stats[l])}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const thStyle = { textAlign: "left", padding: "10px 14px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", color: "#5C6470", borderBottom: `1px solid ${LINE}`, whiteSpace: "nowrap" };
const tdStyle = { padding: "9px 14px", borderBottom: `1px solid ${LINE}`, whiteSpace: "nowrap", color: "#D8DCE2" };

// ---------------------------------------------------------------------------
// Import des stats complètes de match (box score) — nouvel onglet
// ---------------------------------------------------------------------------

function BoxScoreTab({ roster, index, onImported, onDelete }) {
  const [preview, setPreview] = useState(null);
  const [fileErr, setFileErr] = useState("");
  const [date, setDate] = useState(todayLocal());
  const [opponent, setOpponent] = useState("");
  const [opponentScore, setOpponentScore] = useState("");
  const [teamName, setTeamName] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [requestedIds, setRequestedIds] = useState(new Set());
  const fileRef = useRef();

  async function handleDelete(m) {
    await onDelete(m.id, `${m.date} vs ${m.opponent}`);
    setRequestedIds(s => new Set([...s, m.id]));
    setConfirmDeleteId(null);
  }

  async function openFile(id) {
    const f = await storeGet("boxscore_file:" + id);
    if (!f) return;
    const a = document.createElement("a");
    a.href = f.dataUrl;
    a.download = f.name;
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  useEffect(() => { storeGet("team_name").then(v => { if (v) setTeamName(v); }); }, []);
  async function saveTeamName(v) { setTeamName(v); await storeSet("team_name", v); }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileErr(""); setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseBoxScoreFile(buf, roster, teamName);
      if (!parsed.matchedCount) throw new Error("No row could be matched to a roster player (the 1st column must contain the player's name).");
      const fileDataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      setPreview({ ...parsed, fileName: file.name, fileDataUrl });
    } catch (err) { setFileErr(err.message || "Error reading the file."); }
  }

  async function confirmImport() {
    if (!opponent.trim()) { setFileErr("Enter the opponent before confirming."); return; }
    setBusy(true);
    await onImported(preview, { date, opponent: opponent.trim(), opponentScore: opponentScore ? Number(opponentScore) : null });
    setBusy(false); setPreview(null); setOpponent(""); setOpponentScore("");
    if (fileRef.current) fileRef.current.value = "";
  }

  return (
    <div>
      <SectionTitle eyebrow="03 — Official Stats" title="Full match stats" />
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 22, marginBottom: 24 }}>
        <p style={{ color: "#8B93A1", fontSize: 13.5, lineHeight: 1.6, margin: "0 0 16px" }}>
          Drop the full match stats Excel file here (box score: points, rebounds, assists, minutes…).
          Expected format: one row per player, player name in the 1st column, one column per stat.
          These stats complement — not replace — the ones calculated from the coding file.
        </p>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
          <div><label style={labelStyle}>Match date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 180 }} /></div>
          <div style={{ flex: 1, minWidth: 200 }}><label style={labelStyle}>Opponent</label><input type="text" placeholder="e.g. Zalgiris" value={opponent} onChange={e => setOpponent(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ width: 140 }}>
            <label style={labelStyle}>Opponent score</label>
            <input type="number" placeholder="e.g. 78" value={opponentScore} onChange={e => setOpponentScore(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>Exact team name as it appears in the file</label>
          <input type="text" placeholder="e.g. Swiss National Team" value={teamName} onChange={e => saveTeamName(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320 }} />
          <div style={{ fontSize: 11.5, color: "#5C6470", marginTop: 6 }}>
            Needed to identify the correct totals row in the file (especially if the file also lists the opponent's stats) — saved once for all future imports.
          </div>
        </div>
        <div style={{ fontSize: 11.5, color: "#5C6470", marginBottom: 14 }}>The opponent's score is optional, but needed to calculate DRTG in the Team tab.</div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ color: "#8B93A1", fontSize: 13 }} />
        {fileErr && <div style={{ color: RED, fontSize: 13, marginTop: 10 }}>{fileErr}</div>}
        {preview && (
          <div style={{ marginTop: 18, padding: 16, background: PANEL2, borderRadius: 10, border: `1px solid ${LINE}` }}>
            <div style={{ fontSize: 13, color: PAPER, marginBottom: 6 }}>
              Sheet read: <b>{preview.sheetName}</b> · <b>{preview.matchedCount}</b> players recognized ·
              stats detected: {preview.statLabels.join(", ") || "none"}
            </div>
            <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 10 }}>
              Players recognized: {preview.matchedPlayers.join(", ") || "none"}
            </div>
            {preview.teamRow && (
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 10, background: PANEL, border: `1px solid ${preview.teamRowConfirmed ? TEAL : AMBER}`, borderRadius: 8, marginBottom: 12 }}>
                <ClipboardList size={15} color={preview.teamRowConfirmed ? TEAL : AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 12.5, color: "#D8DCE2", lineHeight: 1.5 }}>
                  {preview.teamRowConfirmed ? (
                    <><b>Team totals row confirmed: "{preview.teamRow.label}"</b> (exact name entered above) — used as the source of truth for team stats.</>
                  ) : (
                    <><b>Team totals row assumed: "{preview.teamRow.label}"</b> — only one non-player row found, so likely correct, but not confirmed by the exact name. Enter the team name above to be sure.</>
                  )}
                </div>
              </div>
            )}
            {preview.unmatchedRows.length > 0 && (
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 10, background: PANEL, border: `1px solid ${RED}`, borderRadius: 8, marginBottom: 12 }}>
                <AlertTriangle size={15} color={RED} style={{ marginTop: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 12.5, color: "#D8DCE2", lineHeight: 1.5 }}>
                  <b>{preview.unmatchedRows.length} row(s) not matched to a roster player</b> — excluded from totals:
                  <br />{preview.unmatchedRows.join(", ")}
                  <br />If one of them is actually a player (different spelling, last name only...), let me know so I can improve recognition.
                </div>
              </div>
            )}
            <button disabled={busy} onClick={confirmImport} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "Import…" : "Confirm import"}</button>
          </div>
        )}
      </div>

      <SectionTitle eyebrow="History" title="Box scores imported" />
      {index.length === 0 ? <EmptyState text="No box score imported yet." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {index.slice().reverse().map(m => (
            <div key={m.id} style={{ ...btnRow, cursor: "default" }}>
              <span>{m.date} <span style={{ color: "#5C6470" }}>vs</span> {m.opponent} <span style={{ color: "#5C6470" }}>· {m.matchedCount} players</span></span>
              {requestedIds.has(m.id) ? (
                <span style={{ fontSize: 11.5, color: AMBER }}>Pending admin approval</span>
              ) : confirmDeleteId === m.id ? (
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: 11.5, color: RED }}>Delete this box score?</span>
                  <button onClick={() => handleDelete(m)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                  <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                </div>
              ) : (
                <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                  <button onClick={() => openFile(m.id)} title="Open the original file" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Download size={15} /></button>
                  <button onClick={() => setConfirmDeleteId(m.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={15} /></button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Training — saisie manuelle
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Work plan / objectifs — par thématique, stat ciblée vs stat actuelle
// ---------------------------------------------------------------------------

// Construit la liste des stats "de coding" reliables à un objectif (fréquence, PPPP,
// % Ouvert par playtype/play, attaque et défense) — en plus des stats de box score.
// BUG RÉEL CORRIGÉ (signalé par l'utilisateur) : seules les catégories "Playtypes" et "Plays"
// étaient proposées comme stats liables pour un objectif — même côté défense. Ça excluait
// toute la catégorie "Shot selection" (ex. "% tir contesté"), "Results & misc." (ex. pertes
// de balle provoquées), "Screen defense", "Spacing", et toute catégorie personnalisée —
// rendant impossible de fixer un objectif défensif basé sur ces stats. On parcourt maintenant
// TOUTES les catégories (les mêmes que celles affichées dans Offense/Defense), les mêmes deux
// côtés (attaque/défense), pour que n'importe quel tag configuré dans Settings soit liable.
function buildCodingStatOptions(off, def, cats) {
  const c = cats || currentTagCategories();
  const out = {};
  const BUILTIN_CATEGORIES = ["Playtypes", "Plays", "Shot selection", "Defensive mistakes", "Screen defense", "Spacing", "Shot zone", "Results & misc."];
  const customCategoryNames = Object.keys(c).filter(n => n !== "Player" && !BUILTIN_CATEGORIES.includes(n));
  const allCategoryNames = [...BUILTIN_CATEGORIES, ...customCategoryNames];

  const addSide = (plays, sideLabel) => {
    for (const categoryName of allCategoryNames) {
      const tags = categoryTags(categoryName, c);
      if (!tags.length) continue;
      // BUG RÉEL CORRIGÉ (signalé par l'utilisateur) : "Shot selection" (% Contested/Open)
      // utilisait un calcul générique — % par rapport au total des actions taguées
      // Contested/Open — au lieu de la vraie formule officielle du classeur (shootingSelection),
      // déjà utilisée pour l'affichage "Shooting Selection (defense)" sur la fiche du joueur.
      // Cette formule utilise un dénominateur différent (toutes les actions avec un playtype,
      // pas seulement celles taguées Contested/Open) et regroupe aussi les pertes de balle
      // dans "Contested" — donnant un nombre différent de mon calcul générique. On réutilise
      // maintenant exactement la même fonction, pour que les deux affichages correspondent.
      if (categoryName === "Shot selection") {
        const selection = shootingSelection(plays, c);
        const total = selection.reduce((s, d) => s + d.value, 0);
        selection.forEach(d => {
          out[`[${sideLabel}] Shot selection — ${d.name} (Frequency)`] = total ? (100 * d.value) / total : 0;
        });
        continue;
      }
      // BUG RÉEL CORRIGÉ (signalé par l'utilisateur) : un tag jamais utilisé par le joueur
      // (0 occurrence, ex. "Cut" en attaque) n'apparaissait pas du tout dans la liste des
      // stats liables — "groupBreakdown" (utilisée en interne) exclut délibérément les tags à
      // 0 occurrence, pratique pour l'affichage normal (éviter des lignes vides dans la
      // répartition à l'écran) mais ça empêchait justement de fixer un objectif "faire
      // progresser depuis 0". On calcule donc ici SANS cette exclusion, pour que tous les
      // tags configurés restent sélectionnables, même à 0%.
      const allEntries = tags.map(label => {
        const matching = plays.filter(p => tagIsSet(p.tags, label));
        const pts = matching.reduce((s, p) => s + playPoints(p.tags), 0);
        return { name: label, count: matching.length, pppp: matching.length ? pts / matching.length : 0, open: openPct(matching) };
      });
      const categoryTotal = allEntries.reduce((s, e) => s + e.count, 0);
      allEntries.forEach(item => {
        out[`[${sideLabel}] ${categoryName} — ${item.name} (Frequency)`] = categoryTotal ? (100 * item.count) / categoryTotal : 0;
        if (item.pppp !== null && item.pppp !== undefined) out[`[${sideLabel}] ${categoryName} — ${item.name} (PPPP)`] = item.pppp;
        if (item.open !== null && item.open !== undefined) out[`[${sideLabel}] ${categoryName} — ${item.name} (% Open)`] = item.open;
      });
    }
  };
  addSide(off, "Offense");
  addSide(def, "Defense");
  return out;
}

function useObjectives(playerName) {
  const [objectives, setObjectives] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [playerName]);

  async function load() {
    setLoading(true);
    const raw = (await storeGet("objectives:" + playerName)) || [];
    // BUG RÉEL CORRIGÉ : le format des clés de stats liées au fichier de coding a changé
    // (ancien : "[Attaque] PNR — Fréquence", nouveau : "[Offense] Playtypes — PNR
    // (Frequency)", pour couvrir toutes les catégories et pas seulement Playtypes/Plays) —
    // les objectifs déjà enregistrés avec l'ancien format ne retrouvaient plus leur valeur
    // actuelle, l'app affichait "–" pour eux. On migre silencieusement au chargement.
    const cats = currentTagCategories();
    let changed = false;
    const migrated = raw.map(o => {
      if (!o.linkedStat) return o;
      const m = o.linkedStat.match(/^\[(Attaque|Offense|Defense)\] (.+) — (Fréquence|Frequency|PPPP|% Ouvert|% Open)$/);
      if (!m) return o;
      const [, oldSide, tagName, oldMetric] = m;
      const newSide = oldSide === "Attaque" ? "Offense" : oldSide;
      const metricLabel = { "Fréquence": "Frequency", "Frequency": "Frequency", "PPPP": "PPPP", "% Ouvert": "% Open", "% Open": "% Open" }[oldMetric];
      const categoryName = ["Playtypes", "Plays"].find(cat => categoryTags(cat, cats).some(t => normTag(t) === normTag(tagName)));
      if (!categoryName) return o; // tag introuvable dans ces deux catégories, on ne migre pas à l'aveugle
      const newKey = `[${newSide}] ${categoryName} — ${tagName} (${metricLabel})`;
      changed = true;
      return { ...o, linkedStat: newKey };
    });
    if (changed) await storeSet("objectives:" + playerName, migrated);
    setObjectives(migrated);
    setLoading(false);

    // Deuxième migration, ciblée : la catégorie "Shot selection" a d'abord utilisé le calcul
    // générique (clés du type "[Defense] Shot selection — Open (% Open)" ou "Contested (% Open)"),
    // avant d'être remplacée par la vraie formule officielle du classeur (shootingSelection),
    // qui n'expose plus que "(Frequency)" avec le nom "Contested / Turnover" (pas "Contested"
    // seul). Les objectifs créés entre-temps pointaient vers une clé qui n'existe plus — l'app
    // affichait "–" pour eux, même après la première migration ci-dessus (qui ne couvre que
    // l'ANCIEN format pré-catégories, pas celui-ci).
    const migrated2 = migrated.map(o => {
      if (!o.linkedStat) return o;
      const m2 = o.linkedStat.match(/^\[(Offense|Defense)\] Shot selection — (Open|Contested)(?:\s*\/\s*Turnover)? \([^)]+\)$/);
      if (!m2) return o;
      const [, side, tagName] = m2;
      const newTagName = tagName === "Contested" ? "Contested / Turnover" : "Open";
      return { ...o, linkedStat: `[${side}] Shot selection — ${newTagName} (Frequency)` };
    });
    if (JSON.stringify(migrated2) !== JSON.stringify(migrated)) {
      await storeSet("objectives:" + playerName, migrated2);
      setObjectives(migrated2);
    }
  }

  async function save(obj) {
    const id = obj.id || uid();
    const record = { ...obj, id };
    const next = objectives.some(o => o.id === id) ? objectives.map(o => o.id === id ? record : o) : [...objectives, record];
    await storeSet("objectives:" + playerName, next);
    setObjectives(next);
  }

  async function remove(id) {
    const next = objectives.filter(o => o.id !== id);
    await storeSet("objectives:" + playerName, next);
    setObjectives(next);
  }

  return { objectives, loading, save, remove };
}

function ObjectiveForm({ initial, linkableStats, onSave, onCancel, busy }) {
  const [description, setDescription] = useState(initial?.description || "");
  const [direction, setDirection] = useState(initial?.direction || "up");
  const [linkedStat, setLinkedStat] = useState(initial?.linkedStat || "");
  const [targetValue, setTargetValue] = useState(initial?.targetValue ?? "");
  const [manualCurrent, setManualCurrent] = useState(initial?.manualCurrent ?? "");
  const [startValue, setStartValue] = useState(initial?.startValue ?? "");
  const startDate = initial?.startDate || todayLocal();

  // À la création d'un objectif, si on choisit une stat liée, on capture sa valeur actuelle
  // comme point de départ (point fixe, ne bouge plus ensuite même si la stat évolue).
  function handleLinkedStatChange(l) {
    setLinkedStat(l);
    if (!initial && l && linkableStats[l] !== undefined) setStartValue(Number(linkableStats[l].toFixed(2)));
  }

  const boxKeys = Object.keys(linkableStats).filter(k => !k.startsWith("["));
  const codingKeys = Object.keys(linkableStats).filter(k => k.startsWith("["));

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <div style={{ flex: 1, minWidth: 220 }}>
          <label style={labelStyle}>Objective (description)</label>
          <input value={description} onChange={e => setDescription(e.target.value)} placeholder="e.g. Increase % Open on PnR" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
        </div>
        <div style={{ width: 160 }}>
          <label style={labelStyle}>Desired direction</label>
          <select value={direction} onChange={e => setDirection(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
            <option value="up">Increase ↑</option>
            <option value="down">Decrease ↓</option>
          </select>
        </div>
      </div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <div style={{ width: 320 }}>
          <label style={labelStyle}>Linked stat (updates automatically)</label>
          <select value={linkedStat} onChange={e => handleLinkedStatChange(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
            <option value="">— Manual entry —</option>
            {boxKeys.length > 0 && (
              <optgroup label="Box score">
                {boxKeys.map(l => <option key={l} value={l}>{l}</option>)}
              </optgroup>
            )}
            {codingKeys.length > 0 && (
              <optgroup label="Coding file (playtypes / plays)">
                {codingKeys.map(l => <option key={l} value={l}>{l.replace(/^\[[^\]]+\]\s*/, "")} {l.match(/^\[([^\]]+)\]/)?.[0]}</option>)}
              </optgroup>
            )}
          </select>
        </div>
        <div style={{ width: 130 }}>
          <label style={labelStyle}>Starting value</label>
          <input type="number" value={startValue} onChange={e => setStartValue(e.target.value)} placeholder="starting value" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
        </div>
        <div style={{ width: 150 }}>
          <label style={labelStyle}>Target stat</label>
          <input type="number" value={targetValue} onChange={e => setTargetValue(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
        </div>
        {!linkedStat && (
          <div style={{ width: 150 }}>
            <label style={labelStyle}>Current stat (manual)</label>
            <input type="number" value={manualCurrent} onChange={e => setManualCurrent(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
        )}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onCancel} style={btnSecondary}>Cancel</button>
        <button disabled={busy || !description.trim() || targetValue === ""} onClick={() => onSave({
          id: initial?.id, description, direction, linkedStat, startDate,
          startValue: startValue === "" ? null : Number(startValue),
          targetValue: Number(targetValue), manualCurrent: manualCurrent === "" ? null : Number(manualCurrent),
        })} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "…" : "Save"}</button>
      </div>
    </div>
  );
}

function ObjectiveTrack({ startValue, currentValue, targetValue, progressPct, color }) {
  const pct = Math.max(2, Math.min(98, progressPct ?? 0));
  const fmt = (v) => v === null || v === undefined ? "–" : (Number.isInteger(v) ? v : v.toFixed(1));
  return (
    <div style={{ position: "relative", height: 56, margin: "18px 4px 26px" }}>
      <div style={{ position: "absolute", top: 24, left: 0, right: 0, height: 6, borderRadius: 6, background: "#1B2028" }} />
      <div style={{ position: "absolute", top: 24, left: 0, width: `${pct}%`, height: 6, borderRadius: 6, background: `linear-gradient(90deg, ${color}55, ${color})`, transition: "width 0.4s ease" }} />

      <div style={{ position: "absolute", top: 19, left: 0, width: 14, height: 14, borderRadius: "50%", background: "#5C6470", border: `2px solid ${INK}`, transform: "translateX(-50%)" }} />
      <div style={{ position: "absolute", top: 34, left: 0, fontSize: 10, color: "#5C6470", transform: "translateX(-50%)", whiteSpace: "nowrap" }}>Départ · {fmt(startValue)}</div>

      <div style={{ position: "absolute", top: 16, left: "100%", transform: "translateX(-50%)", fontSize: 15 }}>🏁</div>
      <div style={{ position: "absolute", top: 34, left: "100%", fontSize: 10, color: "#8B93A1", transform: "translateX(-50%)", whiteSpace: "nowrap" }}>Objective · {fmt(targetValue)}</div>

      <div style={{ position: "absolute", top: 15, left: `${pct}%`, width: 22, height: 22, borderRadius: "50%", background: color, border: `3px solid ${INK}`, transform: "translateX(-50%)", boxShadow: `0 0 12px ${color}99` }} />
      <div style={{ position: "absolute", top: -8, left: `${pct}%`, transform: "translateX(-50%)", background: color, color: "#0B1410", fontSize: 11, fontWeight: 800, padding: "2px 9px", borderRadius: 7, whiteSpace: "nowrap" }}>
        {fmt(currentValue)}
      </div>
    </div>
  );
}

function ObjectiveCard({ objective, currentValue, isCoach, onEdit, onDelete }) {
  const hasCurrent = currentValue !== null && currentValue !== undefined;
  const hasStart = objective.startValue !== null && objective.startValue !== undefined;
  let progressPct = null, onTrack = null, delta = null;

  if (hasCurrent && objective.targetValue !== undefined) {
    onTrack = objective.direction === "down" ? currentValue <= objective.targetValue : currentValue >= objective.targetValue;
    if (hasStart) {
      // Progression mesurée depuis le point de départ, pas depuis zéro.
      const range = objective.targetValue - objective.startValue;
      progressPct = range !== 0 ? Math.max(0, Math.min(100, (100 * (currentValue - objective.startValue)) / range)) : (onTrack ? 100 : 0);
      delta = currentValue - objective.startValue;
    } else {
      progressPct = objective.direction === "down"
        ? Math.max(0, Math.min(100, 100 * (objective.targetValue / Math.max(currentValue, 0.0001))))
        : Math.max(0, Math.min(100, (100 * currentValue) / objective.targetValue));
    }
  }
  const barColor = progressPct === null ? "#5C6470" : onTrack ? TEAL : progressPct >= 66 ? AMBER : RED;

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 14, marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 8 }}>
        <div style={{ fontSize: 13.5, fontWeight: 600 }}>{objective.description}</div>
        {isCoach && (
          <div style={{ display: "flex", gap: 10, flexShrink: 0 }}>
            <button onClick={onEdit} style={{ fontSize: 11.5, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>Edit</button>
            <button onClick={onDelete} style={{ fontSize: 11.5, color: RED, background: "none", border: "none", cursor: "pointer" }}>Suppr.</button>
          </div>
        )}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        {delta !== null && (
          <div style={{ fontSize: 12, color: (objective.direction === "down" ? delta <= 0 : delta >= 0) ? TEAL : RED, fontWeight: 700 }}>
            {delta >= 0 ? "+" : ""}{delta.toFixed(1)} depuis le départ
          </div>
        )}
        <div style={{ fontSize: 11, color: "#5C6470" }}>{objective.startDate ? `depuis le ${objective.startDate}` : ""} · {objective.direction === "down" ? "goal: decrease" : "goal: increase"}{objective.linkedStat ? ` · linked to "${objective.linkedStat}"` : " · manual entry"}</div>
      </div>
      <ObjectiveTrack startValue={objective.startValue} currentValue={currentValue} targetValue={objective.targetValue} progressPct={progressPct} color={barColor === "#5C6470" ? AMBER : barColor} />
      {progressPct !== null && <div style={{ fontSize: 10.5, color: "#5C6470", marginTop: -14 }}>{progressPct.toFixed(0)}% du chemin parcouru vers l'objectif</div>}
    </div>
  );
}

function ObjectivesPanel({ playerName, isCoach, box, off, def }) {
  const { objectives, loading, save, remove } = useObjectives(playerName);
  const [editing, setEditing] = useState(null);
  const [busy, setBusy] = useState(false);

  const codingStats = useMemo(() => buildCodingStatOptions(off, def), [off, def]);
  const linkableStats = useMemo(() => ({ ...box.averages, ...codingStats }), [box.averages, codingStats]);

  if (loading) return <EmptyState text="Loading…" />;

  function currentValueFor(obj) {
    if (obj.linkedStat) { const v = linkableStats[obj.linkedStat]; return v !== undefined ? v : null; }
    return obj.manualCurrent;
  }

  return (
    <div>
      {isCoach && !editing && (
        <button onClick={() => setEditing("new")} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: AMBER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", marginBottom: 20 }}>
          <Plus size={14} /> Add an objective
        </button>
      )}
      {editing && (
        <ObjectiveForm
          initial={editing === "new" ? null : editing}
          linkableStats={linkableStats}
          busy={busy}
          onCancel={() => setEditing(null)}
          onSave={async (obj) => { setBusy(true); await save(obj); setBusy(false); setEditing(null); }}
        />
      )}
      {objectives.length === 0 && !editing ? (
        <EmptyState text="No objective defined for this player." />
      ) : (
        objectives.map(o => (
          <ObjectiveCard key={o.id} objective={o} currentValue={currentValueFor(o)} isCoach={isCoach}
            onEdit={() => setEditing(o)} onDelete={() => remove(o.id)} />
        ))
      )}
    </div>
  );
}

const DEFAULT_TRAINING_THEMES = ["Pts + Indiv", "Pts - Indiv", "Coll Off", "Coll Def"];
const TRAINING_THEME_COLORS = { "Pts + Indiv": AMBER, "Pts - Indiv": "#C97BE0", "Coll Off": TEAL, "Coll Def": "#7C9CF2" };
const TRAINING_THEME_FALLBACK_COLORS = ["#E4231C", "#4A90D9", "#B15FE0", "#8B93A1", "#F2A93B", "#2FBF9C"];
function trainingThemeColor(name, allThemes) {
  if (TRAINING_THEME_COLORS[name]) return TRAINING_THEME_COLORS[name];
  const idx = Math.max(0, allThemes.indexOf(name));
  return TRAINING_THEME_FALLBACK_COLORS[idx % TRAINING_THEME_FALLBACK_COLORS.length];
}
function useTrainingThemes() {
  const [themes, setThemes] = useState(DEFAULT_TRAINING_THEMES);
  useEffect(() => { storeGet("training_themes").then(t => setThemes(t && t.length ? t : DEFAULT_TRAINING_THEMES)); }, []);
  async function addTheme(name) {
    const trimmed = name.trim();
    if (!trimmed || themes.includes(trimmed)) return;
    const next = [...themes, trimmed];
    await storeSet("training_themes", next);
    setThemes(next);
  }
  return { themes, addTheme };
}

function useTrainingPlan(playerName) {
  const [plan, setPlan] = useState(null);
  useEffect(() => { storeGet("training_plan:" + playerName).then(p => setPlan(p || { "Pts + Indiv": 25, "Pts - Indiv": 25, "Coll Off": 25, "Coll Def": 25 })); }, [playerName]);
  async function save(p) { await storeSet("training_plan:" + playerName, p); setPlan(p); }
  return { plan, save };
}

function TrainingPlanEditor({ plan, onSave, isCoach }) {
  const [values, setValues] = useState(plan);
  const [busy, setBusy] = useState(false);
  const { themes } = useTrainingThemes();
  useEffect(() => setValues(plan), [plan]);
  const total = themes.reduce((s, t) => s + (Number(values[t]) || 0), 0);

  if (!isCoach) return null;
  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>Target plan — desired session breakdown</div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
        {themes.map(t => (
          <div key={t} style={{ width: 130 }}>
            <label style={{ ...labelStyle, color: trainingThemeColor(t, themes) }}>{t}</label>
            <input type="number" min={0} max={100} value={values[t] || 0} onChange={e => setValues(v => ({ ...v, [t]: Number(e.target.value) }))}
              style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <span style={{ fontSize: 12, color: total === 100 ? TEAL : AMBER }}>{total}% of plan allocated{total !== 100 ? " (ideally 100%)" : ""}</span>
        <button disabled={busy} onClick={async () => { setBusy(true); await onSave(values); setBusy(false); }} style={{ ...btnPrimary, width: "auto", padding: "8px 16px" }}>{busy ? "…" : "Save plan"}</button>
      </div>
    </div>
  );
}

function TrainingLog({ playerName, isCoach }) {
  const [entries, setEntries] = useState([]);
  const { themes, addTheme } = useTrainingThemes();
  const [editingId, setEditingId] = useState(null);
  const [newThemeName, setNewThemeName] = useState("");
  const [form, setForm] = useState({ date: todayLocal(), thematique: themes[0], theme: "", objectif: "", commentaire: "", eval: 3, duree: 15 });
  const [busy, setBusy] = useState(false);
  const { plan, save: savePlan } = useTrainingPlan(playerName);
  const formRef = useRef();

  useEffect(() => { load(); }, [playerName]);
  async function load() { setEntries((await storeGet("training:" + playerName)) || []); }

  function resetForm() {
    setEditingId(null);
    setForm({ date: todayLocal(), thematique: themes[0], theme: "", objectif: "", commentaire: "", eval: 3, duree: 15 });
  }
  function startEdit(e) {
    setEditingId(e.id);
    setForm({ date: e.date, thematique: e.thematique, theme: e.theme || "", objectif: e.objectif || "", commentaire: e.commentaire || "", eval: e.eval ?? null, duree: e.duree ?? 15 });
    if (formRef.current) formRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function save() {
    setBusy(true);
    const next = editingId
      ? entries.map(e => e.id === editingId ? { ...e, ...form } : e)
      : [{ id: uid(), ...form }, ...entries];
    await storeSet("training:" + playerName, next);
    setEntries(next);
    resetForm();
    setBusy(false);
  }
  async function remove(id) {
    const next = entries.filter(e => e.id !== id);
    await storeSet("training:" + playerName, next);
    setEntries(next);
    if (editingId === id) resetForm();
  }

  const total = entries.length;
  const ratedEntries = entries.filter(e => e.eval !== null && e.eval !== undefined);
  const avgEval = ratedEntries.length ? ratedEntries.reduce((s, e) => s + Number(e.eval), 0) / ratedEntries.length : null;
  const totalDuree = entries.reduce((s, e) => s + (Number(e.duree) || 0), 0);
  const dureeByTheme = (t) => entries.filter(e => e.thematique === t).reduce((s, e) => s + (Number(e.duree) || 0), 0);
  const realDistribution = themes.map(t => ({
    name: t, value: dureeByTheme(t), color: trainingThemeColor(t, themes),
  })).filter(d => d.value > 0);
  const realPct = themes.map(t => ({ theme: t, real: totalDuree ? (100 * dureeByTheme(t)) / totalDuree : 0, target: plan ? Number(plan[t]) || 0 : 0 }));

  return (
    <div>
      <div style={{ fontSize: 11, color: "#5C6470", marginBottom: 12 }}>Looking up sessions saved under: "{playerName}"</div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
        <StatPill label="Sessions completed" value={total} />
        <StatPill label="Average rating" value={avgEval !== null ? avgEval.toFixed(1) + "/5" : "–"} sub={avgEval !== null ? undefined : "no session rated"} tone={avgEval !== null ? undefined : "teal"} />
      </div>
      {avgEval !== null && (
        <div style={{ marginBottom: 20 }}>
          <div style={{ width: "100%", maxWidth: 320, height: 10, borderRadius: 6, background: "linear-gradient(90deg, #E4231C, #F0883E, #FFC107, #8BC34A, #2D6A1F)", position: "relative" }}>
            <div style={{ position: "absolute", top: -6, left: `${((avgEval - 1) / 4) * 100}%`, transform: "translateX(-50%)", width: 4, height: 22, borderRadius: 2, background: "#fff", border: `1px solid ${INK}` }} />
          </div>
        </div>
      )}

      {total > 0 && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Rating distribution</div>
          <div style={{ display: "flex", gap: 14, alignItems: "flex-end", height: 90 }}>
            {[1, 2, 3, 4, 5].map(v => {
              const count = entries.filter(e => Number(e.eval) === v).length;
              const maxCount = Math.max(1, ...[1, 2, 3, 4, 5].map(x => entries.filter(e => Number(e.eval) === x).length));
              return (
                <div key={v} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, flex: 1 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: ratingColor(v) }}>{count}</div>
                  <div style={{ width: "100%", height: 50, display: "flex", alignItems: "flex-end" }}>
                    <div style={{ width: "100%", height: `${(count / maxCount) * 100}%`, minHeight: count > 0 ? 4 : 0, borderRadius: "4px 4px 0 0", background: ratingColor(v) }} />
                  </div>
                  <div style={{ fontSize: 11, color: "#5C6470" }}>Rating {v}</div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <TrainingPlanEditor plan={plan || { "Pts + Indiv": 25, "Pts - Indiv": 25, "Coll Off": 25, "Coll Def": 25 }} onSave={savePlan} isCoach={isCoach} />

      {total > 0 && (
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 20 }}>
          <DonutCard title="Actual breakdown (by duration)" data={realDistribution} note="No session with a category and duration." unit="min" />
          <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, flex: "1 1 320px" }}>
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Target vs actual plan</div>
            {realPct.map(r => (
              <div key={r.theme} style={{ marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                  <span style={{ color: TRAINING_THEME_COLORS[r.theme] }}>{r.theme}</span>
                  <span style={{ color: "#8B93A1" }}>{r.real.toFixed(0)}% actual · {r.target}% target</span>
                </div>
                <div style={{ position: "relative", height: 8, borderRadius: 6, background: "#1B2028" }}>
                  <div style={{ position: "absolute", top: 0, left: 0, height: 8, borderRadius: 6, width: `${Math.min(100, r.real)}%`, background: TRAINING_THEME_COLORS[r.theme] }} />
                  <div style={{ position: "absolute", top: -3, left: `${Math.min(100, r.target)}%`, width: 2, height: 14, background: "#fff" }} />
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {isCoach && (
        <div ref={formRef} style={{ background: PANEL, border: `1px solid ${editingId ? AMBER : LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 700 }}>{editingId ? "Edit session" : "New session"}</div>
            {editingId && <button onClick={resetForm} style={{ fontSize: 11.5, color: "#8B93A1", background: "none", border: "none", cursor: "pointer" }}>Cancel edit</button>}
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
            <div><label style={labelStyle}>Date</label><input type="date" value={form.date} onChange={e => setForm(f => ({ ...f, date: e.target.value }))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160 }} /></div>
            <div style={{ width: 180 }}>
              <label style={labelStyle}>Category</label>
              <select value={form.thematique} onChange={e => setForm(f => ({ ...f, thematique: e.target.value }))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: trainingThemeColor(form.thematique, themes), fontWeight: 700 }}>
                {themes.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
              <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                <input value={newThemeName} onChange={e => setNewThemeName(e.target.value)} placeholder="New category…" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", fontSize: 11.5, padding: "6px 8px" }} />
                <button type="button" onClick={async () => { await addTheme(newThemeName); setNewThemeName(""); }} style={{ ...btnSecondary, padding: "6px 10px", fontSize: 11.5 }}>Add</button>
              </div>
            </div>
            <div style={{ flex: 1, minWidth: 160 }}><label style={labelStyle}>Specific theme</label><input value={form.theme} onChange={e => setForm(f => ({ ...f, theme: e.target.value }))} placeholder="e.g. Shot off a PnR" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
            <div style={{ width: 110 }}>
              <label style={labelStyle}>Duration (min)</label>
              <input type="number" min={0} value={form.duree} onChange={e => setForm(f => ({ ...f, duree: Number(e.target.value) }))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
            </div>
            <div style={{ width: 130 }}>
              <label style={labelStyle}>Rating /5</label>
              <select value={form.eval ?? ""} onChange={e => setForm(f => ({ ...f, eval: e.target.value === "" ? null : Number(e.target.value) }))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: form.eval ? ratingColor(form.eval) : "#5C6470", fontWeight: 700 }}>
                <option value="">No rating</option>
                {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            </div>
          </div>
          <div style={{ marginBottom: 10 }}><label style={labelStyle}>Objective</label><input value={form.objectif} onChange={e => setForm(f => ({ ...f, objectif: e.target.value }))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ marginBottom: 12 }}><label style={labelStyle}>Comments</label><textarea value={form.commentaire} onChange={e => setForm(f => ({ ...f, commentaire: e.target.value }))} rows={2} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} /></div>
          <button disabled={busy} onClick={save} style={{ ...btnPrimary, width: "auto", padding: "9px 18px", display: "flex", alignItems: "center", gap: 6 }}>{editingId ? "Save changes" : (<><Plus size={14} /> Add the session</>)}</button>
        </div>
      )}
      {entries.length === 0 ? <EmptyState text="No session recorded." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {entries.map(e => (
            <div key={e.id} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 14, position: "relative" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ fontSize: 13, color: "#5C6470" }}>{e.date}</div>
                  {e.thematique && <span style={{ fontSize: 10.5, fontWeight: 700, color: trainingThemeColor(e.thematique, themes), border: `1px solid ${trainingThemeColor(e.thematique, themes)}`, borderRadius: 5, padding: "1px 6px" }}>{e.thematique}</span>}
                  {e.duree ? <span style={{ fontSize: 11, color: "#5C6470" }}>{e.duree} min</span> : null}
                </div>
                <div style={{ fontFamily: "ui-monospace, monospace", color: ratingColor(e.eval), fontWeight: 700 }}>{e.eval ?? "–"}/5</div>
              </div>
              <div style={{ fontWeight: 600, marginTop: 4 }}>{e.theme}</div>
              {e.objectif && <div style={{ fontSize: 13, color: "#8B93A1", marginTop: 2 }}>Objective: {e.objectif}</div>}
              {e.commentaire && <div style={{ fontSize: 13, color: "#8B93A1", marginTop: 4 }}>{e.commentaire}</div>}
              {isCoach && (
                <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 10 }}>
                  <button onClick={() => startEdit(e)} title="Edit" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><ClipboardList size={14} /></button>
                  <button onClick={() => remove(e.id)} title="Delete" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><X size={14} /></button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Mental evaluation
// ---------------------------------------------------------------------------

// Les 10 critères tels que présents dans le classeur original (WHO life skills).
const MENTAL_CRITERIA = [
  "Problem Solving", "Decision Making", "Creative Thinking", "Critical Thinking",
  "Effective Communication", "Interpersonal Skills", "Self-Awareness", "Empathy",
  "Managing Stress", "Managing Emotions",
];

// ---------------------------------------------------------------------------
// Visibility config — per-team switches controlling what players can see
// (tabs, training, wellness charts...). Coaches always see everything;
// this only ever restricts the "player" role. Managed from the Admin panel.
// ---------------------------------------------------------------------------

const DEFAULT_VISIBILITY = {
  tabs: { players: true, team: true, scouting: true, planning: true },
  playerDetail: { stats: true, objectives: true, training: true, mental: true, wellness: true, role: true, meetings: true },
  team: { standings: true, teamPlay: true, advanced: true, resources: true },
  wellnessCharts: false, // les graphiques Wellness sont cachés aux joueurs par défaut
};

function useVisibilityConfig() {
  const [config, setConfig] = useState(DEFAULT_VISIBILITY);
  const [loading, setLoading] = useState(true);
  useEffect(() => {
    storeGet("visibility_config").then(v => {
      setConfig(v ? { ...DEFAULT_VISIBILITY, ...v, tabs: { ...DEFAULT_VISIBILITY.tabs, ...(v.tabs || {}) }, playerDetail: { ...DEFAULT_VISIBILITY.playerDetail, ...(v.playerDetail || {}) }, team: { ...DEFAULT_VISIBILITY.team, ...(v.team || {}) } } : DEFAULT_VISIBILITY);
      setLoading(false);
    });
  }, []);
  return { config, loading };
}

// ---------------------------------------------------------------------------
// Wellness — 3x/day player check-in (physical + mental), charts hidden from
// players unless explicitly allowed via Admin visibility settings.
// ---------------------------------------------------------------------------

const WELLNESS_SLOTS = [
  { key: "wake", label: "Wake up" },
  { key: "training1", label: "After 1st training" },
  { key: "training2", label: "After 2nd training" },
];
// Tri chronologique correct : par date, puis par ORDRE du créneau dans la journée (Wake up
// → After 1st training → After 2nd training) — pas alphabétique sur le nom du créneau, qui
// plaçait "wake" après "training1"/"training2" par erreur ("t" < "w").
function wellnessSortKey(entry) {
  const idx = WELLNESS_SLOTS.findIndex(s => s.key === entry.slot);
  return entry.date + "-" + String(idx >= 0 ? idx : 9).padStart(2, "0");
}

// ---------------------------------------------------------------------------
// Role — définie par le coach pour chaque joueur : un nom de rôle, une description de ce
// qui est attendu concrètement, et une image le représentant. Le joueur peut la consulter
// mais pas la modifier.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Meeting — journal des entretiens avec un joueur : date, titre, description. Ajouté
// par le coach, consultable par le joueur.
// ---------------------------------------------------------------------------

function MeetingsTab({ playerName, isCoach }) {
  const [meetings, setMeetings] = useState([]);
  const [loading, setLoading] = useState(true);
  const [date, setDate] = useState(todayLocal());
  const [title, setTitle] = useState("");
  const [description, setDescription] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  useEffect(() => { load(); }, [playerName]);
  async function load() {
    setLoading(true);
    setMeetings(((await storeGet("meetings:" + playerName)) || []).sort((a, b) => b.date.localeCompare(a.date)));
    setLoading(false);
  }

  async function addMeeting() {
    if (!title.trim()) return;
    setBusy(true);
    const entry = { id: uid(), date, title: title.trim(), description: description.trim() };
    const next = [entry, ...meetings];
    await storeSet("meetings:" + playerName, next);
    setMeetings(next);
    setTitle(""); setDescription("");
    setBusy(false);
  }

  async function removeMeeting(id) {
    const next = meetings.filter(m => m.id !== id);
    await storeSet("meetings:" + playerName, next);
    setMeetings(next);
    setConfirmDeleteId(null);
  }

  if (loading) return <EmptyState text="Loading…" />;

  return (
    <div>
      {isCoach && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
            <div style={{ width: 160 }}>
              <label style={labelStyle}>Date</label>
              <input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
            </div>
            <div style={{ flex: 1, minWidth: 200 }}>
              <label style={labelStyle}>Title</label>
              <input value={title} onChange={e => setTitle(e.target.value)} placeholder="e.g. Mid-season check-in" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
            </div>
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>Description</label>
            <textarea value={description} onChange={e => setDescription(e.target.value)} rows={4} placeholder="What was discussed, agreed, or decided..."
              style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} />
          </div>
          <button disabled={busy || !title.trim()} onClick={addMeeting} style={{ ...btnPrimary, width: "auto", padding: "9px 18px" }}>{busy ? "…" : "Add meeting"}</button>
        </div>
      )}

      {meetings.length === 0 ? <EmptyState text="No meeting recorded yet." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {meetings.map(m => (
            <div key={m.id} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 6 }}>
                <div>
                  <div style={{ fontSize: 11, color: "#5C6470" }}>{m.date}</div>
                  <div style={{ fontSize: 14.5, fontWeight: 700 }}>{m.title}</div>
                </div>
                {isCoach && (
                  confirmDeleteId === m.id ? (
                    <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                      <button onClick={() => removeMeeting(m.id)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                      <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                    </div>
                  ) : (
                    <button onClick={() => setConfirmDeleteId(m.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex", flexShrink: 0 }}><Trash2 size={14} /></button>
                  )
                )}
              </div>
              {m.description && <div style={{ fontSize: 13, color: "#8B93A1", lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{m.description}</div>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function RoleTab({ playerName, isCoach }) {
  const [role, setRole] = useState(null);
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [image, setImage] = useState("");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const imgRef = useRef();

  useEffect(() => {
    storeGet("role:" + playerName).then(r => {
      setRole(r || null);
      setName((r && r.name) || "");
      setDescription((r && r.description) || "");
      setImage((r && r.image) || "");
    });
  }, [playerName]);

  async function handleImage(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImage(await fileToResizedDataURL(file, 500, 0.85));
  }

  async function save() {
    setBusy(true);
    const record = { name: name.trim(), description: description.trim(), image };
    await storeSet("role:" + playerName, record);
    setRole(record);
    setStatus("Saved.");
    setBusy(false);
  }

  if (role === null && !isCoach) return <EmptyState text="No role defined yet." />;

  if (!isCoach) {
    // Lecture seule pour le joueur.
    return (
      <div>
        {image && <img src={image} alt="" style={{ width: "100%", maxWidth: 460, borderRadius: 14, marginBottom: 18, display: "block" }} />}
        <h2 style={{ fontSize: 18, fontWeight: 700, margin: "0 0 10px" }}>{name || "No role defined yet."}</h2>
        {description && <p style={{ color: "#8B93A1", fontSize: 14, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{description}</p>}
      </div>
    );
  }

  return (
    <div>
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18 }}>
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>Role name</label>
          <input value={name} onChange={e => setName(e.target.value)} placeholder="e.g. Floor general" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>What's expected, concretely</label>
          <textarea value={description} onChange={e => setDescription(e.target.value)} rows={6} placeholder="e.g. Push the pace in transition, take care of the ball, organize the offense..."
            style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} />
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={labelStyle}>Image representing this role</label>
          {image ? (
            <div>
              <img src={image} alt="" style={{ width: "100%", maxWidth: 460, borderRadius: 14, marginBottom: 10, display: "block" }} />
              <div style={{ display: "flex", gap: 10 }}>
                <button type="button" onClick={() => imgRef.current && imgRef.current.click()} style={btnSecondary}>Change image</button>
                <button type="button" onClick={() => setImage("")} style={{ ...btnSecondary, color: RED, borderColor: RED }}>Remove image</button>
              </div>
            </div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <button type="button" onClick={() => imgRef.current && imgRef.current.click()} style={{ width: 100, height: 100, borderRadius: 12, background: PANEL2, border: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0, cursor: "pointer", padding: 0 }}>
                <Camera size={22} color="#5C6470" />
              </button>
              <span style={{ fontSize: 12, color: "#5C6470" }}>Click to choose an image</span>
            </div>
          )}
          <input ref={imgRef} type="file" accept="image/*" onChange={handleImage} style={{ display: "none" }} />
        </div>
        <button disabled={busy} onClick={save} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "…" : "Save"}</button>
        {status && <div style={{ fontSize: 12, color: TEAL, marginTop: 10 }}>{status}</div>}
      </div>
    </div>
  );
}


function WellnessTab({ playerName, isCoach, canSeeCharts, teamId, teamName }) {
  const [entries, setEntries] = useState([]);
  const [date, setDate] = useState(todayLocal());
  const [slot, setSlot] = useState("wake");
  const [physical, setPhysical] = useState(3);
  const [mental, setMental] = useState(3);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [requestedIds, setRequestedIds] = useState(new Set());
  const [filterDate, setFilterDate] = useState(""); // vide = tous les jours

  useEffect(() => { load(); setFilterDate(""); }, [playerName]);
  async function load() { setEntries((await storeGet("wellness:" + playerName)) || []); }

  async function handleDeleteEntry(e) {
    await requestDeletion(teamId, teamName, "wellness", `${playerName} — ${e.date} · ${WELLNESS_SLOTS.find(s => s.key === e.slot)?.label || e.slot}`, { playerName, entryId: e.id });
    setRequestedIds(s => new Set([...s, e.id]));
    setConfirmDeleteId(null);
  }

  async function submit() {
    setBusy(true); setStatus("");
    try {
      // Même correctif que dans HomeTab : relit l'état le plus récent juste avant d'écrire,
      // au lieu de se fier à l'état React local qui peut être périmé si deux soumissions
      // se suivent de près (ex. le coach modifie deux entrées rapidement).
      const latest = (await storeGet("wellness:" + playerName)) || [];
      const existingIdx = latest.findIndex(e => e.date === date && e.slot === slot);
      const entry = { id: existingIdx >= 0 ? latest[existingIdx].id : uid(), date, slot, physical, mental };
      const next = existingIdx >= 0 ? latest.map((e, i) => i === existingIdx ? entry : e) : [entry, ...latest];
      await storeSet("wellness:" + playerName, next);
      setEntries(next);
      setStatus("Saved — thanks!");
    } catch (e) {
      setStatus("Not saved — check your connection and try again.");
    }
    setBusy(false);
  }

  const chartData = [...entries]
    .filter(e => !filterDate || e.date === filterDate)
    .sort((a, b) => wellnessSortKey(a).localeCompare(wellnessSortKey(b)))
    .map(e => ({ label: `${e.date} · ${WELLNESS_SLOTS.find(s => s.key === e.slot)?.label || e.slot}`, physical: e.physical, mental: e.mental }));

  return (
    <div>
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Check-in</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
          <div><label style={labelStyle}>Date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160 }} /></div>
          <div style={{ width: 200 }}>
            <label style={labelStyle}>When</label>
            <select value={slot} onChange={e => setSlot(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
              {WELLNESS_SLOTS.map(s => <option key={s.key} value={s.key}>{s.label}</option>)}
            </select>
          </div>
        </div>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 16 }}>
          <div style={{ width: 220 }}>
            <label style={labelStyle}>How do you feel physically? (1-5)</label>
            <select value={physical} onChange={e => setPhysical(Number(e.target.value))}
              style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: ratingColor(physical), fontWeight: 700 }}>
              {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
            </select>
          </div>
          <div style={{ width: 220 }}>
            <label style={labelStyle}>How do you feel mentally? (1-5)</label>
            <select value={mental} onChange={e => setMental(Number(e.target.value))}
              style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: ratingColor(mental), fontWeight: 700 }}>
              {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
            </select>
          </div>
        </div>
        <button disabled={busy} onClick={submit} style={{ ...btnPrimary, width: "auto", padding: "10px 18px" }}>{busy ? "…" : "Submit"}</button>
        {status && <div style={{ fontSize: 12.5, color: TEAL, marginTop: 10 }}>{status}</div>}
      </div>

      {!canSeeCharts ? (
        !isCoach && <div style={{ fontSize: 12.5, color: "#5C6470" }}>Your check-ins are recorded — trends are visible to your coaching staff.</div>
      ) : entries.length === 0 ? (
        <EmptyState text="No check-in recorded yet." />
      ) : (
        <>
          <div style={{ marginBottom: 16, maxWidth: 220 }}>
            <label style={labelStyle}>View a specific day</label>
            <div style={{ display: "flex", gap: 8 }}>
              <input type="date" value={filterDate} onChange={e => setFilterDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%" }} />
              {filterDate && <button onClick={() => setFilterDate("")} style={{ ...btnSecondary, flexShrink: 0, padding: "9px 12px" }}>All</button>}
            </div>
          </div>
          {filterDate && chartData.length === 0 ? (
            <EmptyState text="No check-in recorded on this day." />
          ) : (
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
            <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "16px 10px 8px 0", flex: "1 1 320px" }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, padding: "0 18px 10px" }}>Physical</div>
              <ResponsiveContainer width="100%" height={200}>
                <ComposedChart data={chartData} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                  <CartesianGrid stroke={LINE} vertical={false} />
                  <XAxis dataKey="label" tick={{ fill: "#5C6470", fontSize: 9 }} axisLine={{ stroke: LINE }} hide={!filterDate} />
                  <YAxis domain={[1, 5]} tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} />
                  <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }} />
                  <Line type="monotone" dataKey="physical" stroke={TEAL} strokeWidth={2.5} dot={{ r: 4, fill: TEAL }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
            <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "16px 10px 8px 0", flex: "1 1 320px" }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: PAPER, padding: "0 18px 10px" }}>Mental</div>
              <ResponsiveContainer width="100%" height={200}>
                <ComposedChart data={chartData} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                  <CartesianGrid stroke={LINE} vertical={false} />
                  <XAxis dataKey="label" tick={{ fill: "#5C6470", fontSize: 9 }} axisLine={{ stroke: LINE }} hide={!filterDate} />
                  <YAxis domain={[1, 5]} tick={{ fill: "#5C6470", fontSize: 10 }} axisLine={{ stroke: LINE }} />
                  <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }} />
                  <Line type="monotone" dataKey="mental" stroke={AMBER} strokeWidth={2.5} dot={{ r: 4, fill: AMBER }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
          )}
        </>
      )}

      {isCoach && entries.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: "#8B93A1", textTransform: "uppercase", marginBottom: 8 }}>Log</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {[...entries].filter(e => !filterDate || e.date === filterDate).sort((a, b) => wellnessSortKey(b).localeCompare(wellnessSortKey(a))).map(e => (
              <div key={e.id} style={{ ...btnRow, cursor: "default" }}>
                <span style={{ fontSize: 12.5 }}>
                  {e.date} <span style={{ color: "#5C6470" }}>· {WELLNESS_SLOTS.find(s => s.key === e.slot)?.label || e.slot}</span>
                  {" — "}Physical <b style={{ color: ratingColor(e.physical) }}>{e.physical}</b> · Mental <b style={{ color: ratingColor(e.mental) }}>{e.mental}</b>
                </span>
                {requestedIds.has(e.id) ? (
                  <span style={{ fontSize: 11.5, color: AMBER }}>Pending admin approval</span>
                ) : confirmDeleteId === e.id ? (
                  <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <span style={{ fontSize: 11.5, color: RED }}>Delete this entry?</span>
                    <button onClick={() => handleDeleteEntry(e)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                    <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                  </div>
                ) : (
                  <button onClick={() => setConfirmDeleteId(e.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={14} /></button>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function MentalLog({ playerName, isCoach }) {
  const [entries, setEntries] = useState([]);
  const [date, setDate] = useState(todayLocal());
  const [comment, setComment] = useState("");
  const [ratings, setRatings] = useState(Object.fromEntries(MENTAL_CRITERIA.map(c => [c, 3])));
  const [busy, setBusy] = useState(false);

  useEffect(() => { load(); }, [playerName]);
  async function load() { setEntries((await storeGet("mental:" + playerName)) || []); }

  async function add() {
    setBusy(true);
    const next = [{ id: uid(), date, ratings, comment }, ...entries];
    await storeSet("mental:" + playerName, next);
    setEntries(next); setComment(""); setBusy(false);
  }
  async function remove(id) {
    const next = entries.filter(e => e.id !== id);
    await storeSet("mental:" + playerName, next);
    setEntries(next);
  }

  // Moyenne par critère sur toutes les évaluations, pour la roue de synthèse.
  const avgRatings = {};
  MENTAL_CRITERIA.forEach(c => {
    const vals = entries.map(e => e.ratings?.[c]).filter(v => v !== undefined && v !== null);
    avgRatings[c] = vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : 3;
  });
  const overallAvg = entries.length ? MENTAL_CRITERIA.reduce((s, c) => s + avgRatings[c], 0) / MENTAL_CRITERIA.length : null;

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
        <StatPill label="Evaluations" value={entries.length} />
        <StatPill label="Overall average" value={overallAvg !== null ? overallAvg.toFixed(1) + "/5" : "–"} />
      </div>

      {entries.length > 0 && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 14, padding: 20, marginBottom: 24, display: "flex", justifyContent: "center" }}>
          <RoseChart ratings={avgRatings} categories={MENTAL_CRITERIA} size={460} emphasizeExtremes={false} />
        </div>
      )}

      {isCoach && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14, alignItems: "flex-end" }}>
            <div><label style={labelStyle}>Date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160 }} /></div>
            <div style={{ flex: 1, minWidth: 200 }}><label style={labelStyle}>Comment</label><input value={comment} onChange={e => setComment(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          </div>
          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", marginBottom: 8 }}>Rating per criterion (1 = weak, 5 = excellent)</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
            {MENTAL_CRITERIA.map(c => (
              <div key={c} style={{ width: 170 }}>
                <label style={{ ...labelStyle, fontSize: 10.5 }}>{c}</label>
                <select value={ratings[c]} onChange={e => setRatings(r => ({ ...r, [c]: Number(e.target.value) }))}
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", padding: "8px 10px", fontSize: 13, color: ratingColor(ratings[c]), fontWeight: 700 }}>
                  {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </div>
            ))}
          </div>
          <button disabled={busy} onClick={add} style={{ ...btnPrimary, width: "auto", padding: "10px 18px" }}>{busy ? "…" : "Add evaluation"}</button>
        </div>
      )}

      {entries.length === 0 ? <EmptyState text="No evaluation recorded." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {entries.map(e => {
            const entryAvg = MENTAL_CRITERIA.reduce((s, c) => s + (e.ratings?.[c] || 0), 0) / MENTAL_CRITERIA.length;
            return (
              <div key={e.id} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: "10px 14px", position: "relative" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div><span style={{ color: "#5C6470", fontSize: 13 }}>{e.date}</span>{e.comment && <span style={{ marginLeft: 10, fontSize: 13.5 }}>{e.comment}</span>}</div>
                  <div style={{ fontFamily: "ui-monospace, monospace", color: ratingColor(Math.round(entryAvg)), fontWeight: 700 }}>{entryAvg.toFixed(1)}/5</div>
                </div>
                {isCoach && <button onClick={() => remove(e.id)} style={{ position: "absolute", top: 10, right: 10, background: "none", border: "none", color: "#5C6470", cursor: "pointer" }}><X size={13} /></button>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Team
// ---------------------------------------------------------------------------

function MatchSelector({ options, selectedKeys, onChange }) {
  const [open, setOpen] = useState(false);
  const allSelected = selectedKeys === null;
  const count = allSelected ? options.length : selectedKeys.size;

  function toggleKey(k) {
    const next = new Set(allSelected ? options.map(o => matchKey(o.date, o.opponent)) : selectedKeys);
    if (next.has(k)) next.delete(k); else next.add(k);
    onChange(next.size === options.length ? null : next);
  }

  if (!options.length) return null;

  return (
    <div style={{ position: "relative", marginBottom: 20 }}>
      <button onClick={() => setOpen(o => !o)} style={{
        display: "flex", alignItems: "center", gap: 8, padding: "9px 14px", background: PANEL, border: `1px solid ${LINE}`,
        borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit",
      }}>
        <ClipboardList size={14} color={AMBER} />
        {allSelected ? `All matches (${options.length})` : `${count} match${count !== 1 ? "es" : ""} selected`}
        <ChevronRight size={13} style={{ transform: open ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
      </button>
      {open && (
        <div style={{ position: "absolute", zIndex: 10, marginTop: 6, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, padding: 10, minWidth: 280, maxHeight: 320, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          <button onClick={() => onChange(null)} style={{ ...btnRow, marginBottom: 6, background: allSelected ? PANEL : "transparent" }}>
            <span style={{ fontWeight: 700 }}>All matches</span>
            {allSelected && <span style={{ color: AMBER }}>✓</span>}
          </button>
          {options.map(o => {
            const k = matchKey(o.date, o.opponent);
            const checked = allSelected || selectedKeys.has(k);
            return (
              <button key={k} onClick={() => toggleKey(k)} style={{ ...btnRow, marginBottom: 4 }}>
                <span>{o.date} <span style={{ color: "#5C6470" }}>vs</span> {o.opponent}
                  <span style={{ color: "#5C6470", fontSize: 11 }}> {o.hasCoding && o.hasBox ? "· coding + box" : o.hasCoding ? "· coding" : "· box"}</span>
                </span>
                {checked && <span style={{ color: AMBER }}>✓</span>}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function TeamAdvancedStats({ advanced }) {
  const STAT_LABELS_FR = { fgm: "Shots made (FGM)", fga: "Shots attempted (FGA)", tpm: "3pt made", fta: "FT attempted", ftm: "FT made", oreb: "Off. rebounds", tov: "Turnovers", pts: "Points" };

  if (advanced.loading) return <EmptyState text="Loading…" />;
  if (!advanced.perMatch.length) {
    return (
      <div>
        <EmptyState text="No box score found in memory. Import a file from the 'Full Stats' tab (top menu) — this file, not the coding one, feeds this tab." />
      </div>
    );
  }

  const avg = (key) => {
    const vals = advanced.perMatch.map(m => m[key]).filter(v => v !== null && v !== undefined);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  };
  const weighted = computeWeightedTeamPercentages(advanced.perMatch);
  const ortg = avg("ortg"), drtg = avg("drtg"), poss = avg("poss"), ftRate = avg("ftRate"), oreb = avg("oreb"), reb = avg("reb");
  const { pct2, pct3, pctFT, efg, tovPct, orebPct, astPct } = weighted;
  const dreb = avg("dreb"), ast = avg("ast"), pts = avg("pts"), ptse = advanced.perMatch.some(m => m.opponentScore !== null && m.opponentScore !== undefined)
    ? advanced.perMatch.map(m => m.opponentScore).filter(v => v !== null && v !== undefined).reduce((s, v, _, arr) => s + v / arr.length, 0) : null;
  const net = ortg !== null && drtg !== null ? ortg - drtg : null;
  const approxPoss = advanced.perMatch.some(m => m.approxPoss);

  return (
    <div>
      <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 12, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, marginBottom: 20 }}>
        <ClipboardList size={15} color={TEAL} style={{ marginTop: 2, flexShrink: 0 }} />
        <div style={{ fontSize: 12.5, color: "#8B93A1", lineHeight: 1.6 }}>
          <b style={{ color: PAPER }}>Columns detected in the imported file(s):</b> {advanced.rawLabels.join(", ") || "none"}.
          <br />
          <b style={{ color: PAPER }}>Columns recognized for calculations:</b>{" "}
          {Object.entries(advanced.columns).filter(([, v]) => v).map(([k, v]) => `${k} → "${v}"`).join(" · ") || "none"}.
        </div>
      </div>

      {advanced.missing && advanced.missing.length > 0 && (
        <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 12, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, marginBottom: 20 }}>
          <AlertTriangle size={15} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
          <div style={{ fontSize: 12.5, color: "#8B93A1", lineHeight: 1.5 }}>
            Still not recognized: {advanced.missing.map(k => STAT_LABELS_FR[k]).join(", ")}.
            Compare with the list of detected columns just above — if the exact name in your file
            doesn't match any expected pattern, give me that exact name and I'll add it to detection.
          </div>
        </div>
      )}

      <SectionTitle eyebrow="Ratings" title="ORTG / DRTG" />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
        <StatPill label="Possessions / game" value={poss !== null ? poss.toFixed(1) : "–"} sub={approxPoss ? "approx. (off. rebounds not isolated)" : undefined} />
        <StatPill label="ORTG" value={ortg !== null ? ortg.toFixed(1) : "–"} sub={approxPoss ? "approx." : "points / 100 possessions"} tone="teal" />
        <StatPill label="DRTG" value={drtg !== null ? drtg.toFixed(1) : "–"} sub="points allowed / 100 poss. — requires the opponent's score" tone="red" />
        <StatPill label="Net Rating" value={net !== null ? (net > 0 ? "+" : "") + net.toFixed(1) : "–"} />
      </div>
      {approxPoss && (
        <div style={{ fontSize: 11.5, color: "#5C6470", marginBottom: 20, marginTop: -10 }}>
          Your file doesn't separate offensive and defensive rebounds (single "Reb" column) — possessions, ORTG and DRTG are
          therefore calculated without the offensive rebound term, which overestimates them slightly. Total rebounds: {reb !== null ? reb.toFixed(1) + "/game" : "not detected"}.
        </div>
      )}

      <SectionTitle eyebrow="Four Factors" title="Offensive factors" />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
        <StatPill label="eFG%" value={efg !== null ? efg.toFixed(1) + "%" : "–"} />
        <StatPill label="TOV%" value={tovPct !== null ? tovPct.toFixed(1) + "%" : "–"} tone="red" />
        <StatPill label="OREB%" value={orebPct !== null ? orebPct.toFixed(1) + "%" : "–"} sub="off. rebounds ÷ (missed shots + 0.44 × missed FT)" />
        <StatPill label="FTA/FGA" value={ftRate !== null ? ftRate.toFixed(2) : "–"} />
      </div>
      <div style={{ fontSize: 11.5, color: "#5C6470", marginBottom: 26 }}>
        Defensive Four Factors (same factors on the opponent's side) can't be calculated: they require the opponent's full box score, which isn't imported today.
      </div>

      <SectionTitle eyebrow="Playmaking" title="Assist %" />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
        <StatPill label="AST%" value={astPct !== null ? astPct.toFixed(1) + "%" : "–"} sub="assists ÷ (made shots + 0.44 × made FT)" tone="teal" />
      </div>

      <SectionTitle eyebrow="Detail" title="Shots, rebounds, team play" />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 26 }}>
        <StatPill label="Points / game" value={pts !== null ? pts.toFixed(1) : "–"} tone="teal" />
        <StatPill label="Points allowed / game" value={ptse !== null ? ptse.toFixed(1) : "–"} sub="requires the opponent's score" tone="red" />
        <StatPill label="% 2pts" value={pct2 !== null ? pct2.toFixed(1) + "%" : "–"} />
        <StatPill label="% 3pts" value={pct3 !== null ? pct3.toFixed(1) + "%" : "–"} />
        <StatPill label="% FT" value={pctFT !== null ? pctFT.toFixed(1) + "%" : "–"} />
        <StatPill label="Off. rebounds / game" value={oreb !== null ? oreb.toFixed(1) : "–"} />
        <StatPill label="Def. rebounds / game" value={dreb !== null ? dreb.toFixed(1) : "–"} sub={dreb === null ? "requires off. rebounds + total" : undefined} />
        <StatPill label="Tot. rebounds / game" value={reb !== null ? reb.toFixed(1) : "–"} />
        <StatPill label="Assists / game" value={ast !== null ? ast.toFixed(1) : "–"} sub={ast === null ? "column not detected" : undefined} />
      </div>

      <SectionTitle eyebrow="Detail" title="Per game" />
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
          <thead>
            <tr>
              <th style={thStyle}>Date</th><th style={thStyle}>Opponent</th><th style={thStyle}>Opp. score</th>
              <th style={thStyle}>Poss.</th><th style={thStyle}>ORTG</th><th style={thStyle}>DRTG</th><th style={thStyle}>eFG%</th><th style={thStyle}>TOV%</th>
            </tr>
          </thead>
          <tbody>
            {advanced.perMatch.map((m, i) => (
              <tr key={i}>
                <td style={tdStyle}>{m.date}</td><td style={tdStyle}>{m.opponent}</td>
                <td style={tdStyle}>{m.opponentScore ?? "–"}</td>
                <td style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{m.poss !== null ? m.poss.toFixed(1) : "–"}</td>
                <td style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{m.ortg !== null ? m.ortg.toFixed(1) : "–"}</td>
                <td style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{m.drtg !== null ? m.drtg.toFixed(1) : "–"}</td>
                <td style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{m.efg !== null ? (m.efg * 100).toFixed(1) + "%" : "–"}</td>
                <td style={{ ...tdStyle, fontFamily: "ui-monospace, monospace" }}>{m.tovPct !== null ? (m.tovPct * 100).toFixed(1) + "%" : "–"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Onglet Scouting
// ---------------------------------------------------------------------------

function useScoutingTeams() {
  const [teams, setTeams] = useState({}); // name -> {stats, source, updatedAt}
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, []);

  async function load() {
    setLoading(true);
    const names = (await storeGet("scouting_teams")) || [];
    const out = {};
    for (const n of names) {
      const t = await storeGet("scouting:" + n);
      if (t) out[n] = t;
    }
    setTeams(out);
    setLoading(false);
  }

  async function saveTeam(name, stats, source) {
    const trimmed = name.trim();
    if (!trimmed) return;
    const names = (await storeGet("scouting_teams")) || [];
    if (!names.includes(trimmed)) { await storeSet("scouting_teams", [...names, trimmed]); }
    const existing = await storeGet("scouting:" + trimmed);
    const record = { stats, source, updatedAt: todayLocal(), logo: existing?.logo };
    await storeSet("scouting:" + trimmed, record);
    setTeams(t => ({ ...t, [trimmed]: record }));
  }

  async function saveLogo(name, logo) {
    const existing = (await storeGet("scouting:" + name)) || {};
    const record = { ...existing, logo };
    await storeSet("scouting:" + name, record);
    setTeams(t => ({ ...t, [name]: record }));
  }

  async function deleteTeam(name) {
    const names = ((await storeGet("scouting_teams")) || []).filter(n => n !== name);
    await storeSet("scouting_teams", names);
    await storeDelete("scouting:" + name);
    setTeams(t => { const c = { ...t }; delete c[name]; return c; });
  }

  return { teams, loading, saveTeam, saveLogo, deleteTeam, reload: load };
}

// Calcule les stats "Our team" dans le même schéma, à partir des box scores déjà importés.
function ourTeamAsScoutStats(advanced, box) {
  const avg = (key) => {
    const vals = advanced.perMatch.map(m => m[key]).filter(v => v !== null && v !== undefined);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : undefined;
  };
  const weighted = computeWeightedTeamPercentages(advanced.perMatch);
  const stats = {
    mj: advanced.perMatch.length || undefined,
    poss: avg("poss"), ortg: avg("ortg"), drtg: avg("drtg"),
    efg: weighted.efg ?? undefined,
    pctbp: weighted.tovPct ?? undefined,
    ftafga: avg("ftRate"),
    pts: avg("pts"),
    ptse: advanced.perMatch.some(m => m.opponentScore !== null && m.opponentScore !== undefined)
      ? advanced.perMatch.map(m => m.opponentScore).filter(v => v !== null && v !== undefined).reduce((s, v, _, arr) => s + v / arr.length, 0) : undefined,
    r2: avg("made2"), t2: (avg("made2") !== undefined && avg("missed2") !== undefined) ? avg("made2") + avg("missed2") : undefined,
    pct2: weighted.pct2 ?? undefined,
    r3: avg("made3"), t3: (avg("made3") !== undefined && avg("missed3") !== undefined) ? avg("made3") + avg("missed3") : undefined,
    pct3: weighted.pct3 ?? undefined,
    lfr: avg("madeFT"), lft: (avg("madeFT") !== undefined && avg("missedFT") !== undefined) ? avg("madeFT") + avg("missedFT") : undefined,
    pctlf: weighted.pctFT ?? undefined,
    ro: avg("oreb"), rd: avg("dreb"), rt: avg("reb"),
    pd: avg("ast"), bp: avg("tov"),
  };
  Object.keys(stats).forEach(k => { if (stats[k] === undefined || Number.isNaN(stats[k])) delete stats[k]; });
  return stats;
}

// ---------------------------------------------------------------------------
// Scouting Report — collectif (vidéo + plan de match) et individuel (fiches joueur
// façon "Personnel Scouting" avec roue de niveaux par compétence)
// ---------------------------------------------------------------------------

const RATING_LEVELS = [
  { value: 1, label: "Weak", color: "#E4231C" },
  { value: 2, label: "Not comfortable", color: "#F0883E" },
  { value: 3, label: "Sometimes used", color: "#FFC107" },
  { value: 4, label: "Under control", color: "#8BC34A" },
  { value: 5, label: "Master", color: "#2D6A1F" },
];
const SKILL_CATEGORIES = ["Finishing", "Shooting", "Drive", "Perimeter defense", "Inside defense", "Off ball", "Rebound", "Speed", "Force", "Passing"];
const CLOSEOUT_LEVELS = [
  { key: "easy", label: "Control the Drive", color: TEAL },
  { key: "medium", label: "Medium", color: AMBER },
  { key: "hard", label: "Hard", color: RED },
];
function ratingColor(v) { return v === null || v === undefined ? "#5C6470" : (RATING_LEVELS.find(l => l.value === v) || RATING_LEVELS[2]).color; }

// Roue à secteurs colorés (façon PDF) — pas un radar classique à polygone : chaque
// compétence est un secteur indépendant, sa couleur ET sa taille reflètent la note 1-5.
function RoseChart({ ratings, categories = SKILL_CATEGORIES, size = 420, emphasizeExtremes = true }) {
  const n = categories.length;
  const cx = size / 2, cy = size / 2;
  const maxR = size / 2 - 130; // marge généreuse pour que les libellés (ex. "Perimeter defense") ne débordent jamais
  const angleStep = (2 * Math.PI) / n;

  const sorted = [...categories].sort((a, b) => (ratings[b] || 1) - (ratings[a] || 1));
  const topStrengths = emphasizeExtremes ? new Set(sorted.slice(0, 2)) : new Set();
  const worstWeakness = emphasizeExtremes ? sorted[sorted.length - 1] : null;
  // BUG RÉEL CORRIGÉ : quand deux catégories mises en avant (force/faiblesse, texte agrandi)
  // sont adjacentes sur la roue, leurs libellés se chevauchaient. Un premier correctif les
  // éloignait bien trop (l'un se retrouvait quasiment à l'opposé de la roue) — on utilise
  // maintenant un léger décalage angulaire (les deux mots s'écartent un peu l'un de l'autre,
  // tangentiellement) combiné à un tout petit décalage de rayon, pour qu'ils restent proches
  // de la roue sans jamais se toucher. BUG RÉEL CORRIGÉ : ce décalage était en pixels fixes,
  // calibré pour une roue de 480px — à une taille plus petite (ex. 420px pour l'export en
  // page pleine), le chevauchement revenait. Le décalage est maintenant proportionnel à la
  // taille réelle de la roue, pour rester efficace à n'importe quelle taille.
  const sizeScale = size / 480;
  const emphasizedIndices = categories.map((cat, i) => ({ cat, i, emphasized: topStrengths.has(cat) || cat === worstWeakness })).filter(x => x.emphasized);
  const staggerExtra = {}, angleNudge = {};
  emphasizedIndices.forEach((entry, idx) => {
    const conflict = emphasizedIndices.slice(0, idx).find(other => {
      const diff = Math.min(Math.abs(entry.i - other.i), n - Math.abs(entry.i - other.i));
      return diff <= 2; // catégories voisines ou quasi-voisines sur la roue
    });
    if (conflict) {
      staggerExtra[entry.cat] = 34 * sizeScale;
      // S'écarte du côté opposé à l'autre libellé en conflit, le long du cercle.
      const forward = (entry.i - conflict.i + n) % n <= n / 2;
      angleNudge[entry.cat] = forward ? 0.28 : -0.28;
    } else {
      staggerExtra[entry.cat] = 14 * sizeScale;
      angleNudge[entry.cat] = 0;
    }
  });

  function wedgePath(i, value) {
    const r = Math.max(14, (maxR * value) / 5);
    const a0 = -Math.PI / 2 + i * angleStep + 0.02;
    const a1 = -Math.PI / 2 + (i + 1) * angleStep - 0.02;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    return `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 0 1 ${x1} ${y1} Z`;
  }
  function labelPos(i, extra, nudge = 0) {
    const a = -Math.PI / 2 + (i + 0.5) * angleStep + nudge;
    const r = maxR + 30 + extra;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  }

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ overflow: "visible" }}>
      <circle cx={cx} cy={cy} r={maxR} fill="#0C0F14" />
      {categories.map((cat, i) => (
        <path key={cat} d={wedgePath(i, ratings[cat] || 1)} fill={ratingColor(ratings[cat] || 1)} stroke="#0C0F14" strokeWidth={2} />
      ))}
      {categories.map((cat, i) => {
        const isStrength = topStrengths.has(cat);
        const isWeakness = cat === worstWeakness;
        const emphasis = isStrength || isWeakness;
        const fontSize = emphasis ? 26 : 11.5;
        const color = isStrength ? "#3DDC6F" : isWeakness ? "#FF5C4D" : "#C9CFD8";
        const weight = emphasis ? 800 : 400;
        const words = cat.split(" ");
        const p = labelPos(i, emphasis ? staggerExtra[cat] : 0, emphasis ? angleNudge[cat] : 0);
        return (
          <text key={cat} x={p.x} y={p.y} fill={color} fontSize={fontSize} fontWeight={weight} textAnchor="middle" dominantBaseline="middle" style={{ fontFamily: "Inter, sans-serif" }}>
            {words.map((w, wi) => (
              <tspan key={wi} x={p.x} dy={wi === 0 ? -((words.length - 1) * (fontSize * 0.55)) : fontSize * 1.1}>{w}</tspan>
            ))}
          </text>
        );
      })}
    </svg>
  );
}

function CloseOutBadge({ level }) {
  const l = CLOSEOUT_LEVELS.find(c => c.key === level) || CLOSEOUT_LEVELS[1];
  return <span style={{ color: l.color, fontWeight: 800, fontStyle: "italic", fontSize: 22 }}>{l.label}</span>;
}

function useScoutingReport(teamName) {
  const [collective, setCollective] = useState({ videoUrl: "", notes: "" });
  const [players, setPlayers] = useState([]); // [{id, ...}]
  const [loading, setLoading] = useState(true);

  useEffect(() => { if (teamName) load(); }, [teamName]);

  async function load() {
    setLoading(true);
    const c = await storeGet("scouting_report_collective:" + teamName);
    setCollective(c || { videoUrl: "", notes: "" });
    const ids = (await storeGet("scouting_players:" + teamName)) || [];
    const out = [];
    for (const id of ids) {
      const p = await storeGet("scouting_player:" + teamName + ":" + id);
      if (!p) continue;
      // Migration silencieuse : "Force" a été traduit par erreur en "Strength" ("force
      // physique" n'a pas le même sens que "Strength" en anglais) — les notes déjà données
      // sous l'ancien nom sont reprises sous le bon, pour ne rien perdre.
      if (p.ratings && p.ratings.Strength !== undefined && p.ratings.Force === undefined) {
        const { Strength, ...rest } = p.ratings;
        p.ratings = { ...rest, Force: Strength };
        await storeSet("scouting_player:" + teamName + ":" + id, p);
      }
      out.push(p);
    }
    setPlayers(out);
    setLoading(false);
  }

  async function saveCollective(c) { await storeSet("scouting_report_collective:" + teamName, c); setCollective(c); }

  async function savePlayer(player) {
    const id = player.id || uid();
    const record = { ...player, id };
    await storeSet("scouting_player:" + teamName + ":" + id, record);
    const ids = (await storeGet("scouting_players:" + teamName)) || [];
    if (!ids.includes(id)) await storeSet("scouting_players:" + teamName, [...ids, id]);
    setPlayers(ps => { const i = ps.findIndex(p => p.id === id); if (i === -1) return [...ps, record]; const c = [...ps]; c[i] = record; return c; });
  }

  async function deletePlayer(id) {
    await storeDelete("scouting_player:" + teamName + ":" + id);
    const ids = ((await storeGet("scouting_players:" + teamName)) || []).filter(x => x !== id);
    await storeSet("scouting_players:" + teamName, ids);
    setPlayers(ps => ps.filter(p => p.id !== id));
  }

  // Réordonne les joueurs scoutés selon l'envie du coach — l'ordre est celui du tableau
  // d'identifiants, donc on n'a qu'à échanger deux positions adjacentes puis le persister.
  async function movePlayer(id, direction) {
    const ids = (await storeGet("scouting_players:" + teamName)) || [];
    const i = ids.indexOf(id);
    if (i === -1) return;
    const j = direction === "up" ? i - 1 : i + 1;
    if (j < 0 || j >= ids.length) return; // déjà en haut/en bas, rien à faire
    const next = [...ids];
    [next[i], next[j]] = [next[j], next[i]];
    await storeSet("scouting_players:" + teamName, next);
    setPlayers(ps => {
      const byId = Object.fromEntries(ps.map(p => [p.id, p]));
      return next.map(pid => byId[pid]).filter(Boolean);
    });
  }

  return { collective, players, loading, saveCollective, savePlayer, deletePlayer, movePlayer };
}

function embedUrl(url) {
  if (!url) return null;
  const yt = url.match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|embed\/))([\w-]+)/);
  if (yt) return `https://www.youtube.com/embed/${yt[1]}`;
  if (/drive\.google\.com/.test(url)) return url.replace("/view", "/preview");
  return url; // lien direct (mp4, Vimeo déjà en /embed, etc.)
}

function ScoutingCollective({ collective, onSave, isCoach }) {
  const [videoUrl, setVideoUrl] = useState(collective.videoUrl || "");
  const [notes, setNotes] = useState(collective.notes || "");
  const [busy, setBusy] = useState(false);
  const [embedFailed, setEmbedFailed] = useState(false);
  const [localVideo, setLocalVideo] = useState(null); // {url, name} — session uniquement, jamais sauvegardé
  const fileRef = useRef();
  useEffect(() => { setVideoUrl(collective.videoUrl || ""); setNotes(collective.notes || ""); setEmbedFailed(false); }, [collective.videoUrl, collective.notes]);
  useEffect(() => () => { if (localVideo) URL.revokeObjectURL(localVideo.url); }, [localVideo]); // libère la mémoire au démontage
  const embed = embedUrl(collective.videoUrl);

  function handleLocalFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (localVideo) URL.revokeObjectURL(localVideo.url);
    setLocalVideo({ url: URL.createObjectURL(file), name: file.name });
  }

  return (
    <div>
      {isCoach && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 12, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 10, marginBottom: 16 }}>
            <AlertTriangle size={15} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
            <div style={{ fontSize: 12, color: "#8B93A1", lineHeight: 1.5 }}>
              The file import below plays the video right away in your browser, but is <b>not saved</b> —
              you'll need to re-import it on every visit, and only you can see it. The link, on the other hand, is saved and shared with the whole team.
            </div>
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>Import a local video file (plays immediately, not saved)</label>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <input ref={fileRef} type="file" accept="video/*" onChange={handleLocalFile} style={{ color: "#8B93A1", fontSize: 13 }} />
              {localVideo && <span style={{ fontSize: 12, color: TEAL }}>{localVideo.name} ✓</span>}
            </div>
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>— OR — Video link (YouTube, Google Drive, Vimeo, direct mp4…), saved and shared</label>
            <input value={videoUrl} onChange={e => setVideoUrl(e.target.value)} placeholder="https://…" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>Game plan / team notes</label>
            <textarea value={notes} onChange={e => setNotes(e.target.value)} rows={5} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} />
          </div>
          <button disabled={busy} onClick={async () => { setBusy(true); await onSave({ videoUrl, notes }); setBusy(false); }} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>
            {busy ? "…" : "Save link + notes"}
          </button>
        </div>
      )}

      {localVideo && (
        <div style={{ borderRadius: 12, overflow: "hidden", border: `1px solid ${LINE}`, marginBottom: 20, background: "#000" }}>
          <video src={localVideo.url} controls style={{ width: "100%", display: "block", maxHeight: 480 }} />
        </div>
      )}

      {embed ? (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div style={{ fontSize: 12.5, color: "#8B93A1" }}>If the video doesn't show below (blocked by the environment), open it directly:</div>
            <a href={collective.videoUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12.5, color: AMBER, textDecoration: "none", whiteSpace: "nowrap", marginLeft: 10 }}>Open the video ↗</a>
          </div>
          {!embedFailed && (
            <div style={{ position: "relative", paddingBottom: "56.25%", height: 0, borderRadius: 12, overflow: "hidden", border: `1px solid ${LINE}`, marginBottom: 20, background: PANEL }}>
              <iframe src={embed} style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", border: "none" }}
                allow="autoplay; fullscreen" allowFullScreen title="Scouting video" onError={() => setEmbedFailed(true)} />
            </div>
          )}
        </>
      ) : (
        !localVideo && <EmptyState text="No video added yet." />
      )}

      {collective.notes && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, whiteSpace: "pre-wrap", fontSize: 14, lineHeight: 1.6, color: "#D8DCE2" }}>
          {collective.notes}
        </div>
      )}
    </div>
  );
}

function ScoutingPlayerForm({ initial, onSave, onCancel, busy }) {
  const [firstName, setFirstName] = useState(initial?.firstName || "");
  const [lastName, setLastName] = useState(initial?.lastName || "");
  const [jersey, setJersey] = useState(initial?.jersey || "");
  const [position, setPosition] = useState(initial?.position || "");
  const [plan, setPlan] = useState(initial?.plan || "");
  const [closeOut, setCloseOut] = useState(initial?.closeOut || "medium");
  const [height, setHeight] = useState(initial?.height || "");
  const [handedness, setHandedness] = useState(initial?.handedness || "");
  const [ratings, setRatings] = useState(initial?.ratings || Object.fromEntries(SKILL_CATEGORIES.map(c => [c, 3])));
  const [highlights, setHighlights] = useState(initial?.highlights || []);
  const [photo, setPhoto] = useState(initial?.photo || null);
  const [photoErr, setPhotoErr] = useState("");
  const [photoBusy, setPhotoBusy] = useState(false);
  const fileRef = useRef();

  async function handlePhoto(e) {
    const file = e.target.files[0];
    if (!file) return;
    setPhotoErr(""); setPhotoBusy(true);
    try {
      const url = await fileToResizedDataURL(file, 300, 0.85);
      setPhoto(url);
    } catch (err) {
      setPhotoErr("Photo illisible — essaie un autre fichier (JPG/PNG).");
    }
    setPhotoBusy(false);
    e.target.value = "";
  }

  function addHighlight() { setHighlights(h => [...h, { label: "", value: "" }]); }
  function updateHighlight(i, field, v) { setHighlights(h => h.map((x, j) => j === i ? { ...x, [field]: v } : x)); }
  function removeHighlight(i) { setHighlights(h => h.filter((_, j) => j !== i)); }

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ display: "flex", gap: 16, marginBottom: 14, flexWrap: "wrap" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8 }}>
          <div style={{ width: 72, height: 72, borderRadius: 10, background: PANEL2, border: `1px solid ${LINE}`, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {photo ? <img src={photo} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <Camera size={20} color="#5C6470" />}
          </div>
          <button type="button" onClick={() => fileRef.current && fileRef.current.click()} style={{ fontSize: 11, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>{photoBusy ? "…" : "Photo"}</button>
          <input ref={fileRef} type="file" accept="image/*" onChange={handlePhoto} style={{ display: "none" }} />
          {photoErr && <div style={{ fontSize: 10, color: RED, maxWidth: 90, textAlign: "center" }}>{photoErr}</div>}
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", flex: 1 }}>
          <div style={{ width: 160 }}><label style={labelStyle}>First name</label><input value={firstName} onChange={e => setFirstName(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ width: 160 }}><label style={labelStyle}>Last name</label><input value={lastName} onChange={e => setLastName(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ width: 90 }}><label style={labelStyle}>No.</label><input value={jersey} onChange={e => setJersey(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ width: 140 }}>
            <label style={labelStyle}>Position</label>
            <select value={position} onChange={e => setPosition(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
              <option value="">—</option>
              {["1", "1-2", "2", "2-3", "3", "3-4", "4", "4-5", "5"].map(p => <option key={p} value={p}>Position {p}</option>)}
            </select>
          </div>
          <div style={{ width: 200 }}>
            <label style={labelStyle}>Close Out</label>
            <select value={closeOut} onChange={e => setCloseOut(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
              {CLOSEOUT_LEVELS.map(c => <option key={c.key} value={c.key}>{c.label}</option>)}
            </select>
          </div>
          <div style={{ width: 110 }}>
            <label style={labelStyle}>Height (cm)</label>
            <input type="number" value={height} onChange={e => setHeight(e.target.value)} placeholder="195" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
          </div>
          <div style={{ width: 130 }}>
            <label style={labelStyle}>Dominant hand</label>
            <select value={handedness} onChange={e => setHandedness(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }}>
              <option value="">—</option>
              <option value="Right">Right</option>
              <option value="Left">Left</option>
              <option value="Ambidextrous">Ambidextrous</option>
            </select>
          </div>
        </div>
      </div>

      <div style={{ marginBottom: 14 }}>
        <label style={labelStyle}>Game plan for this player</label>
        <textarea value={plan} onChange={e => setPlan(e.target.value)} rows={3} placeholder="e.g. Push Left - No Reject - Stay engaged" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} />
      </div>

      <div style={{ marginBottom: 14 }}>
        <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", marginBottom: 8 }}>Ratings per skill (1 = Weak, 5 = Master)</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {SKILL_CATEGORIES.map(cat => (
            <div key={cat} style={{ width: 150 }}>
              <label style={{ ...labelStyle, fontSize: 10.5 }}>{cat}</label>
              <select value={ratings[cat] || 3} onChange={e => setRatings(r => ({ ...r, [cat]: Number(e.target.value) }))}
                style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", padding: "8px 10px", fontSize: 13, color: ratingColor(ratings[cat] || 3), fontWeight: 700 }}>
                {RATING_LEVELS.map(l => <option key={l.value} value={l.value}>{l.value} — {l.label}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>

      <div style={{ marginBottom: 14 }}>
        <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", marginBottom: 8 }}>Stats to highlight (e.g. "9.0" / "Pts")</div>
        {highlights.map((h, i) => (
          <div key={i} style={{ marginBottom: 10, padding: "8px 10px", background: PANEL2, borderRadius: 8 }}>
            <div style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center" }}>
              <input value={h.value} onChange={e => updateHighlight(i, "value", e.target.value)} placeholder="9.0 / 40%" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 110 }} />
              <input value={h.label} onChange={e => updateHighlight(i, "label", e.target.value)} placeholder="Pts / % on PnR" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", flex: 1 }} />
              <button onClick={() => removeHighlight(i)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer" }}><X size={14} /></button>
            </div>
            <div style={{ display: "flex", gap: 14, alignItems: "center", fontSize: 11.5, color: "#8B93A1", flexWrap: "wrap" }}>
              <label style={{ display: "flex", gap: 6, alignItems: "center" }}>
                Shape size
                <input type="number" min={40} max={160} step={2} value={h.size ?? 74} onChange={e => updateHighlight(i, "size", Number(e.target.value))}
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 64, padding: "4px 8px", fontSize: 12 }} />
                px
              </label>
              <label style={{ display: "flex", gap: 6, alignItems: "center" }}>
                Text size
                <input type="number" min={0.6} max={2.5} step={0.1} value={h.fontScale ?? 1} onChange={e => updateHighlight(i, "fontScale", Number(e.target.value))}
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 64, padding: "4px 8px", fontSize: 12 }} />
                ×
              </label>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                Color
                {Array.from(new Set([TEAL, ...PLANNING_COLORS])).map(c => (
                  <button key={c} type="button" onClick={() => updateHighlight(i, "color", c)} style={{
                    width: 18, height: 18, borderRadius: "50%", background: c, cursor: "pointer", padding: 0,
                    border: (h.color ?? TEAL) === c ? `2px solid ${PAPER}` : "2px solid transparent",
                  }} />
                ))}
              </div>
            </div>
          </div>
        ))}
        <button onClick={addHighlight} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px dashed ${LINE}`, borderRadius: 8, color: "#8B93A1", fontSize: 12.5, padding: "6px 12px", cursor: "pointer" }}>
          <Plus size={13} /> Add a stat
        </button>
      </div>

      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onCancel} style={btnSecondary}>Cancel</button>
        <button disabled={busy} onClick={() => onSave({ id: initial?.id, firstName, lastName, jersey, position, plan, closeOut, height, handedness, ratings, highlights, photo })}
          style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "…" : "Save"}</button>
      </div>
    </div>
  );
}

const STAT_SHAPES = ["circle", "hexagon", "diamond", "square", "triangle"];
function shapeForLabel(label) {
  let h = 0;
  for (const c of String(label || "")) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return STAT_SHAPES[h % STAT_SHAPES.length];
}
function StatShapeBadge({ label, value, size = 74, fontScale = 1, color = TEAL }) {
  const shape = shapeForLabel(label);
  const c = size / 2;
  const shapes = {
    circle: <circle cx={c} cy={c} r={c - 3} />,
    square: <rect x={5} y={5} width={size - 10} height={size - 10} rx={6} />,
    diamond: <polygon points={`${c},4 ${size - 4},${c} ${c},${size - 4} 4,${c}`} />,
    triangle: <polygon points={`${c},4 ${size - 5},${size - 8} 5,${size - 8}`} />,
    hexagon: <polygon points={`${c},2 ${size - 3},${c * 0.5} ${size - 3},${c * 1.5} ${c},${size - 2} 3,${c * 1.5} 3,${c * 0.5}`} />,
  };
  return (
    <div style={{ position: "relative", width: size, height: size, flexShrink: 0 }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ position: "absolute", inset: 0 }}>
        <g fill={color} stroke="#03231F" strokeWidth={1.5}>{shapes[shape]}</g>
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", textAlign: "center", padding: 6 }}>
        <div style={{ fontWeight: 800, fontSize: 14 * fontScale, color: "#03231F", lineHeight: 1.1 }}>{value}</div>
        <div style={{ fontSize: 8 * fontScale, fontWeight: 700, color: "#03231F", lineHeight: 1.1 }}>{label}</div>
      </div>
    </div>
  );
}

function ScoutingPlayerCard({ player, isCoach, bgPhoto, bgDarkness, teamLogo, onEdit, onDelete, printMode, onMoveUp, onMoveDown, isFirst, isLast }) {
  // 0 = aucun voile (photo visible à 100%), 100 = fond entièrement noir (photo invisible).
  const darkness = Math.max(0, Math.min(100, bgDarkness ?? 70)) / 100;
  // Mode export : la fiche doit remplir toute une page A4 paysage, avec TOUS les éléments
  // bien visibles (photo, stats, chiffres) — pas la version compacte utilisée à l'écran.
  const photoW = printMode ? 220 : 150, photoH = printMode ? 250 : 150;
  const sidebarW = printMode ? 260 : 200;
  const nameSize = printMode ? 23 : 15, subSize = printMode ? 14 : 11.5;
  const chartSize = printMode ? 420 : 445;
  const badgeMinSize = printMode ? 78 : 56;
  return (
    <div data-no-split="true" style={{
      position: "relative",
      background: bgPhoto
        ? `linear-gradient(rgba(22,27,34,${darkness}), rgba(22,27,34,${darkness})), url(${bgPhoto})`
        : PANEL,
      // "100% 100%" (au lieu de "cover") étire la photo pour remplir tout le cadre sans jamais
      // la couper ni la répéter en mosaïque — quitte à légèrement déformer l'image si ses
      // proportions ne correspondent pas exactement à celles du cadre.
      backgroundSize: "100% 100%", backgroundPosition: "center", backgroundRepeat: "no-repeat",
      border: `1px solid ${LINE}`, borderRadius: 14, padding: printMode ? 28 : 16, marginBottom: 14,
    }}>
      {/* Logo de l'équipe scoutée, à l'intérieur du cadre de la fiche (pas en dehors, sur le
          document partagé) — demandé par l'utilisateur, à la même taille que les onglets. */}
      {teamLogo && <img src={teamLogo} alt="" style={{ position: "absolute", top: 16, right: 16, width: 30, height: 30, borderRadius: 6, objectFit: "cover" }} />}
      <div style={{ display: "flex", gap: printMode ? 36 : 24, flexWrap: "wrap", width: "100%" }}>
        <div style={{ width: sidebarW, flexShrink: 0 }}>
          <div style={{ width: photoW, height: photoH, borderRadius: 12, background: PANEL2, border: `1px solid ${LINE}`, overflow: "hidden", marginBottom: printMode ? 12 : 12 }}>
            {player.photo && <img src={player.photo} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />}
          </div>
          <div style={{ fontSize: nameSize, fontWeight: 800, color: PAPER }}>{player.jersey ? `#${player.jersey} ` : ""}{player.lastName?.toUpperCase()}</div>
          <div style={{ fontSize: subSize, color: "#8B93A1", marginBottom: printMode ? 10 : 10 }}>{player.firstName}{player.position ? ` · Position ${player.position}` : ""}</div>

          {(player.height || player.handedness) && (
            <div style={{ display: "flex", gap: 8, marginBottom: printMode ? 10 : 12 }}>
              {player.height && (
                <div style={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, padding: printMode ? "12px 16px" : "8px 12px", flex: 1, textAlign: "center" }}>
                  <div style={{ fontWeight: 800, fontSize: printMode ? 22 : 16, color: PAPER }}>{player.height} cm</div>
                  <div style={{ fontSize: printMode ? 12 : 9.5, color: "#5C6470", textTransform: "uppercase" }}>Height</div>
                </div>
              )}
              {player.handedness && (
                <div style={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, padding: printMode ? "12px 16px" : "8px 12px", flex: 1, textAlign: "center" }}>
                  <div style={{ fontWeight: 800, fontSize: printMode ? 22 : 16, color: PAPER }}>{player.handedness}</div>
                  <div style={{ fontSize: printMode ? 12 : 9.5, color: "#5C6470", textTransform: "uppercase" }}>Dominant hand</div>
                </div>
              )}
            </div>
          )}

          {player.plan && (
            <div style={{ fontSize: printMode ? 13.5 : 12.5, color: "#D8DCE2", marginBottom: printMode ? 10 : 10, lineHeight: 1.4, whiteSpace: "pre-wrap" }}>
              <div style={{ fontSize: printMode ? 13 : 10.5, textTransform: "uppercase", color: "#5C6470", marginBottom: 4 }}>Game plan</div>
              {player.plan}
            </div>
          )}
          <div style={{ fontSize: printMode ? 13.5 : 12.5, marginBottom: printMode ? 10 : 10 }}>Close Out : <CloseOutBadge level={player.closeOut} /></div>
          {player.highlights?.length > 0 && (
            <div style={{ display: "flex", gap: printMode ? 14 : 10, flexWrap: "wrap" }}>
              {player.highlights.map((h, i) => (
                <StatShapeBadge key={i} label={h.label} value={h.value} size={printMode ? Math.max(badgeMinSize, h.size ?? 74) : (h.size ?? 74)} fontScale={(h.fontScale ?? 1) * (printMode ? 1.15 : 1)} color={h.color ?? TEAL} />
              ))}
            </div>
          )}
          {isCoach && !printMode && (
            <div style={{ display: "flex", gap: 10, marginTop: 14, alignItems: "center" }}>
              <button onClick={onEdit} style={{ fontSize: 12, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>Edit</button>
              <button onClick={onDelete} style={{ fontSize: 12, color: RED, background: "none", border: "none", cursor: "pointer" }}>Delete</button>
              {(onMoveUp || onMoveDown) && (
                <div style={{ display: "flex", gap: 4, marginLeft: "auto" }}>
                  <button onClick={onMoveUp} disabled={isFirst} title="Move up" style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: isFirst ? "#3A414C" : "#8B93A1", cursor: isFirst ? "default" : "pointer", padding: 4, display: "flex" }}><ChevronLeft size={14} style={{ transform: "rotate(90deg)" }} /></button>
                  <button onClick={onMoveDown} disabled={isLast} title="Move down" style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: isLast ? "#3A414C" : "#8B93A1", cursor: isLast ? "default" : "pointer", padding: 4, display: "flex" }}><ChevronLeft size={14} style={{ transform: "rotate(-90deg)" }} /></button>
                </div>
              )}
            </div>
          )}
        </div>

        <div style={{ flex: "1 1 420px", display: "flex", justifyContent: "center" }}>
          <RoseChart ratings={player.ratings} size={chartSize} />
        </div>
      </div>
    </div>
  );
}

// Onglet "Staff" du Scouting Report : fichiers déposés par le staff (PDF, images) pour cette
// équipe précise, et reprise automatique des stats/graphiques d'Observation pour la MÊME
// équipe si elle a déjà été observée — pas besoin de redupliquer l'information ailleurs.
function ScoutingStaffPanel({ teamName, isCoach }) {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [observation, setObservation] = useState(null); // { plays, imports } | null
  const fileRef = useRef();

  useEffect(() => { load(); }, [teamName]);
  async function load() {
    setLoading(true);
    setFiles((await storeGet("scouting_staff_files:" + teamName)) || []);
    const allObserved = (await storeGet("scouting_observations")) || {};
    setObservation(allObserved[teamName] || null);
    setLoading(false);
  }

  async function handleUpload(e) {
    const picked = Array.from(e.target.files || []);
    if (!picked.length) return;
    const added = [];
    for (const file of picked) {
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      added.push({ id: uid(), name: file.name, type: file.type, dataUrl, addedAt: new Date().toISOString() });
    }
    const next = [...files, ...added];
    await storeSet("scouting_staff_files:" + teamName, next);
    setFiles(next);
    if (fileRef.current) fileRef.current.value = "";
  }

  async function removeFile(id) {
    const next = files.filter(f => f.id !== id);
    await storeSet("scouting_staff_files:" + teamName, next);
    setFiles(next);
  }

  function openFile(f) {
    const a = document.createElement("a");
    a.href = f.dataUrl; a.download = f.name; a.target = "_blank";
    document.body.appendChild(a); a.click(); a.remove();
  }

  if (loading) return <EmptyState text="Loading…" />;

  const off = (observation?.plays || []).filter(isOffense);
  const def = (observation?.plays || []).filter(isDefense);

  return (
    <div>
      {isCoach && (
        <div style={{ marginBottom: 20 }}>
          <label style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: AMBER, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            <Upload size={14} /> Add PDF or image files
            <input ref={fileRef} type="file" accept="application/pdf,image/*" multiple onChange={handleUpload} style={{ display: "none" }} />
          </label>
        </div>
      )}

      {files.length === 0 ? (
        <EmptyState text="No file shared by the staff yet." />
      ) : (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 12, marginBottom: 28 }}>
          {files.map(f => (
            <div key={f.id} style={{ width: 160, background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 10 }}>
              <button onClick={() => openFile(f)} style={{ width: "100%", height: 90, borderRadius: 6, background: PANEL2, border: "none", cursor: "pointer", overflow: "hidden", marginBottom: 8, padding: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                {f.type?.startsWith("image/") ? <img src={f.dataUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ClipboardList size={26} color="#5C6470" />}
              </button>
              <div style={{ fontSize: 11.5, color: "#D8DCE2", wordBreak: "break-word", marginBottom: 6 }}>{f.name}</div>
              {isCoach && <button onClick={() => removeFile(f.id)} style={{ fontSize: 11, color: RED, background: "none", border: "none", cursor: "pointer" }}>Remove</button>}
            </div>
          ))}
        </div>
      )}

      <SectionTitle eyebrow="Observation" title="Stats & charts for this team" />
      {!observation || (off.length === 0 && def.length === 0) ? (
        <EmptyState text="This team hasn't been observed yet (Observation tab) — its stats and charts will appear here automatically once it has." />
      ) : (
        <OffenseDefenseBreakdown off={off} def={def} categories={currentObservationTagCategories()} />
      )}
    </div>
  );
}

function ScoutingReportTab({ isCoach, teamNames, scoutingTeams, onSaveLogo, initialTeam }) {
  const [selectedTeam, setSelectedTeam] = useState(initialTeam || "");
  const [exportReport, setExportReport] = useState(null);
  const report = useScoutingReport(selectedTeam || null);
  const [subtab, setSubtab] = useState("collectif");
  const [editing, setEditing] = useState(null); // null | "new" | player object
  const [busy, setBusy] = useState(false);
  // Le logo est le MÊME que celui déjà géré dans "Manage teams" (pas une copie séparée) — on
  // le lit directement depuis là, pour que le mettre à jour ici ou là-bas revienne au même.
  const teamLogo = selectedTeam ? scoutingTeams?.[selectedTeam]?.logo : null;
  // La photo de fond des fiches joueurs, elle, est propre à ce nouvel usage (pas de champ
  // équivalent dans "Manage teams") — reste stockée séparément par équipe, avec un réglage de
  // luminosité (0 = aucun voile, la photo est visible à 100% ; 100 = fond entièrement noir).
  const [teamBg, setTeamBg] = useState(null); // { photo, darkness } | null
  useEffect(() => {
    if (!selectedTeam) { setTeamBg(null); return; }
    storeGet("scouting_team_bg:" + selectedTeam).then(v => setTeamBg(v || null));
  }, [selectedTeam]);
  async function handleLogoUpload(e) {
    const file = e.target.files[0];
    if (!file || !selectedTeam) return;
    await onSaveLogo(selectedTeam, await fileToResizedDataURL(file, 200, 0.9));
  }
  async function handleBgUpload(e) {
    const file = e.target.files[0];
    if (!file || !selectedTeam) return;
    const dataUrl = await fileToResizedDataURL(file, 900, 0.85);
    const next = { photo: dataUrl, darkness: teamBg?.darkness ?? 70 };
    await storeSet("scouting_team_bg:" + selectedTeam, next);
    setTeamBg(next);
  }
  async function handleBgDarknessChange(darkness) {
    if (!teamBg || !selectedTeam) return;
    const next = { ...teamBg, darkness };
    setTeamBg(next); // réponse immédiate au curseur, sans attendre l'écriture
    await storeSet("scouting_team_bg:" + selectedTeam, next);
  }

  return (
    <div>
      <div style={{ marginBottom: 20 }}>
        <label style={labelStyle}>Team this report is about</label>
        <select value={selectedTeam} onChange={e => { setSelectedTeam(e.target.value); setEditing(null); }} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320 }}>
          <option value="">— Choose an imported team —</option>
          {teamNames.map(n => <option key={n} value={n}>{n}</option>)}
        </select>
      </div>

      {!selectedTeam ? (
        <EmptyState text="First select an imported team ('Manage teams' tab) to create or view its Scouting Report." />
      ) : (
        <>
      <div style={{ display: "flex", gap: 8, marginBottom: 20, borderBottom: `1px solid ${LINE}`, paddingBottom: 10 }}>
        {[["collectif", "Collective"], ["individuel", "Individual"], ["staff", "Staff"]].map(([id, label]) => (
          <button key={id} onClick={() => setSubtab(id)} style={{
            padding: "7px 12px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "inherit",
            background: subtab === id ? PANEL2 : "transparent", color: subtab === id ? AMBER : "#8B93A1"
          }}>{label}</button>
        ))}
      </div>

      {report.loading ? <EmptyState text="Loading…" /> : (
        <>
          {subtab === "collectif" && <ScoutingCollective collective={report.collective} onSave={report.saveCollective} isCoach={isCoach} />}
          {subtab === "individuel" && (
            <div>
              <div className="screen-only" style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap", justifyContent: "space-between", alignItems: "center" }}>
                {isCoach && !editing ? (
                  <button onClick={() => setEditing("new")} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: AMBER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                    <Plus size={14} /> Add a player
                  </button>
                ) : <div />}
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  {report.players.length > 0 && (
                    <button onClick={async () => {
                      const filename = `scouting_${selectedTeam}_${todayLocal()}.html`;
                      const pdfOk = await tryExportPdf("scouting-print-content", filename, "light");
                      if (pdfOk) return;
                      const r = buildReportHtml("scouting-print-content", filename, "light");
                      if (!r) { alert("Content not found — try again after the page has fully loaded."); return; }
                      tryDownload(r.full, r.filename);
                      setExportReport(r);
                    }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                      <Download size={14} /> Export ({report.players.length} player{report.players.length !== 1 ? "s" : ""}) — HTML → PDF
                    </button>
                  )}
                  {/* Logo de l'équipe scoutée, en haut à droite, à la même taille que les onglets Collective/Individual. */}
                  {isCoach && (
                    <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11, color: "#8B93A1" }}>
                      {teamLogo ? <img src={teamLogo} alt="" style={{ width: 26, height: 26, borderRadius: 6, objectFit: "cover" }} /> : <span style={{ width: 26, height: 26, borderRadius: 6, border: `1px dashed ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center" }}><Camera size={12} /></span>}
                      Team logo
                      <input type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: "none" }} />
                    </label>
                  )}
                  {!isCoach && teamLogo && <img src={teamLogo} alt="" style={{ width: 26, height: 26, borderRadius: 6, objectFit: "cover" }} />}
                  {isCoach && (
                    <>
                      <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11, color: "#8B93A1" }}>
                        <span style={{ width: 26, height: 26, borderRadius: 6, border: `1px dashed ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center", backgroundImage: teamBg?.photo ? `url(${teamBg.photo})` : "none", backgroundSize: "cover" }}>{!teamBg?.photo && <Camera size={12} />}</span>
                        Card background
                        <input type="file" accept="image/*" onChange={handleBgUpload} style={{ display: "none" }} />
                      </label>
                      {teamBg?.photo && (
                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, color: "#8B93A1" }}>
                          Darkness
                          <input type="range" min={0} max={100} value={teamBg.darkness ?? 70} onChange={e => handleBgDarknessChange(Number(e.target.value))} style={{ width: 90 }} />
                          <span style={{ width: 26, textAlign: "right", fontFamily: "ui-monospace, monospace" }}>{teamBg.darkness ?? 70}</span>
                        </label>
                      )}
                    </>
                  )}
                </div>
              </div>
              {editing && (
                <ScoutingPlayerForm
                  initial={editing === "new" ? null : editing}
                  busy={busy}
                  onCancel={() => setEditing(null)}
                  onSave={async (player) => { setBusy(true); await report.savePlayer(player); setBusy(false); setEditing(null); }}
                />
              )}
              {report.players.length === 0 && !editing ? (
                <EmptyState text="No player scouted yet." />
              ) : (
                <div className="screen-only">
                  {report.players.map((p, i) => (
                    <ScoutingPlayerCard key={p.id} player={p} isCoach={isCoach} bgPhoto={teamBg?.photo} bgDarkness={teamBg?.darkness} teamLogo={teamLogo}
                      onEdit={() => setEditing(p)} onDelete={() => report.deletePlayer(p.id)}
                      onMoveUp={() => report.movePlayer(p.id, "up")} onMoveDown={() => report.movePlayer(p.id, "down")}
                      isFirst={i === 0} isLast={i === report.players.length - 1} />
                  ))}
                </div>
              )}
              {report.players.length > 0 && (
                <div className="print-only" id="scouting-print-content">
                  <div style={{ padding: 24, background: "#ffffff", color: "#1A1D24" }}>
                    <h1 style={{ fontSize: 18, marginBottom: 8 }}>Scouting individuel — {selectedTeam}</h1>
                    {report.players.map((p, i) => (
                      <div key={p.id} data-new-page={i % 2 === 0 ? "true" : undefined} style={{ marginBottom: 0, pageBreakInside: "avoid", pageBreakBefore: i % 2 === 0 ? "always" : "auto" }}>
                        <ScoutingPlayerCard player={p} isCoach={false} bgPhoto={teamBg?.photo} bgDarkness={teamBg?.darkness} teamLogo={teamLogo} onEdit={() => {}} onDelete={() => {}} />
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
          {subtab === "staff" && <ScoutingStaffPanel teamName={selectedTeam} isCoach={isCoach} />}
        </>
      )}
        </>
      )}
      <ExportModal report={exportReport} onClose={() => setExportReport(null)} />
    </div>
  );
}

function ScoutingTeamForm({ initial, onSave, onCancel, busy }) {
  const [name, setName] = useState(initial?.name || "");
  const [stats, setStats] = useState(initial?.stats || {});
  const groups = Array.from(new Set(SCOUT_STAT_SCHEMA.map(s => s.group)));

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ marginBottom: 14 }}>
        <label style={labelStyle}>Team name</label>
        <input value={name} onChange={e => setName(e.target.value)} placeholder="ex. Espoirs Cholet" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320 }} />
      </div>
      {groups.map(g => (
        <div key={g} style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", marginBottom: 8 }}>{g}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {SCOUT_STAT_SCHEMA.filter(s => s.group === g).map(s => (
              <div key={s.key} style={{ width: 130 }}>
                <label style={{ ...labelStyle, fontSize: 10.5 }}>{s.label}</label>
                <input type="number" value={stats[s.key] ?? ""} onChange={e => setStats(st => ({ ...st, [s.key]: e.target.value === "" ? undefined : Number(e.target.value) }))}
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", padding: "8px 10px", fontSize: 13 }} />
              </div>
            ))}
          </div>
        </div>
      ))}
      <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
        <button onClick={onCancel} style={btnSecondary}>Cancel</button>
        <button disabled={busy} onClick={() => onSave(name, stats)} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "…" : "Save"}</button>
      </div>
    </div>
  );
}

function ScoutingPhotoImport({ onExtracted }) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setBusy(true); setError("");
    try {
      const teams = await extractScoutingFromImage(file);
      onExtracted(teams);
    } catch (err) { setError(err.message || "Lecture impossible."); }
    setBusy(false);
    if (fileRef.current) fileRef.current.value = "";
  }

  return (
    <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Camera size={16} color={AMBER} />
        <div style={{ fontSize: 13.5, fontWeight: 700 }}>Import from photo</div>
      </div>
      <p style={{ color: "#8B93A1", fontSize: 12.5, lineHeight: 1.5, margin: "0 0 12px" }}>
        Take a screenshot or a photo of a stats table (yours or another team's from the league) — Claude reads it
        and suggests the detected values, which you can correct before saving.
      </p>
      <input ref={fileRef} type="file" accept="image/*" onChange={handleFile} style={{ color: "#8B93A1", fontSize: 13 }} disabled={busy} />
      {busy && <div style={{ fontSize: 12.5, color: AMBER, marginTop: 10 }}>Reading image…</div>}
      {error && <div style={{ color: RED, fontSize: 13, marginTop: 10 }}>{error}</div>}
    </div>
  );
}

function statCellColor(key, valueA, valueB) {
  const schema = SCOUT_STAT_SCHEMA.find(s => s.key === key);
  if (valueA === undefined || valueB === undefined || valueA === valueB) return null;
  const aIsBetter = schema?.lowerBetter ? valueA < valueB : valueA > valueB;
  return aIsBetter ? TEAL : RED;
}

// Rang d'une équipe sur une stat donnée, parmi toutes les équipes chargées dans Scouting
// (comparaisons déjà importées) — sert de proxy pour "le rang dans le championnat".
function leagueRank(teams, key, teamName) {
  const schema = SCOUT_STAT_SCHEMA.find(s => s.key === key);
  const entries = Object.entries(teams)
    .filter(([, t]) => t.stats[key] !== undefined)
    .map(([name, t]) => ({ name, val: t.stats[key] }))
    .sort((x, y) => schema?.lowerBetter ? x.val - y.val : y.val - x.val);
  const idx = entries.findIndex(e => e.name === teamName);
  if (idx === -1 || entries.length < 2) return null;
  return { rank: idx + 1, total: entries.length };
}

function ScoutingComparison({ teams }) {
  const names = Object.keys(teams);
  const [teamA, setTeamA] = useState("");
  const [teamB, setTeamB] = useState("");

  useEffect(() => {
    if (!teamA && names.length) setTeamA(names[0]);
    if (!teamB && names.length > 1) setTeamB(names.find(n => n !== names[0]) || "");
  }, [names.join(",")]);

  if (names.length < 2) return <EmptyState text="Add at least two teams (including yours) to start a comparison." />;

  const a = teams[teamA], b = teams[teamB];
  const radarKeys = ["pts", "ptse", "efg", "pctbp", "pctro", "ortg", "drtg"];
  const radarData = radarKeys.filter(k => a?.stats[k] !== undefined || b?.stats[k] !== undefined).map(k => {
    const schema = SCOUT_STAT_SCHEMA.find(s => s.key === k);
    // Normalisation simple 0-100 pour l'affichage radar (échelle relative entre les 2 équipes).
    const va = a?.stats[k] ?? 0, vb = b?.stats[k] ?? 0;
    const max = Math.max(va, vb, 1);
    return { stat: schema.label, [teamA]: max ? (100 * va) / max : 0, [teamB]: max ? (100 * vb) / max : 0 };
  });

  return (
    <div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
        <div>
          <label style={labelStyle}>Team A</label>
          <select value={teamA} onChange={e => setTeamA(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 220 }}>
            {names.map(n => <option key={n} value={n}>{n}</option>)}
          </select>
        </div>
        <div>
          <label style={labelStyle}>Team B</label>
          <select value={teamB} onChange={e => setTeamB(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 220 }}>
            {names.map(n => <option key={n} value={n}>{n}</option>)}
          </select>
        </div>
      </div>

      <SectionTitle eyebrow="Zoom" title="Four Factors, ratings & % 3pts" />
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 26 }}>
        {["efg", "pctbp", "pctro", "ftafga", "ortg", "drtg", "pct3"].map(key => {
          const schema = SCOUT_STAT_SCHEMA.find(s => s.key === key);
          const va = a?.stats[key], vb = b?.stats[key];
          if (va === undefined && vb === undefined) return null;
          const colorA = statCellColor(key, va, vb), colorB = statCellColor(key, vb, va);
          const rankA = leagueRank(teams, key, teamA), rankB = leagueRank(teams, key, teamB);
          const fmt = (v) => v === undefined ? "–" : (schema.pct ? v.toFixed(1) + "%" : v.toFixed(1));
          return (
            <div key={key} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: "16px 20px", minWidth: 170, flex: "1 1 170px" }}>
              <div style={{ fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", marginBottom: 10 }}>{schema.label}</div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                <div>
                  <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 24, fontWeight: 800, color: colorA || AMBER }}>{fmt(va)}</div>
                  {rankA && <div style={{ fontSize: 10, color: "#5C6470" }}>#{rankA.rank}/{rankA.total} du groupe</div>}
                </div>
                <div style={{ fontSize: 10, color: "#5C6470", maxWidth: 70, textAlign: "right", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{teamA}</div>
              </div>
              <div style={{ height: 1, background: LINE, margin: "8px 0" }} />
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                <div>
                  <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 24, fontWeight: 800, color: colorB || TEAL }}>{fmt(vb)}</div>
                  {rankB && <div style={{ fontSize: 10, color: "#5C6470" }}>#{rankB.rank}/{rankB.total} du groupe</div>}
                </div>
                <div style={{ fontSize: 10, color: "#5C6470", maxWidth: 70, textAlign: "right", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{teamB}</div>
              </div>
            </div>
          );
        })}
      </div>

      {radarData.length > 0 && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 24 }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8 }}>Profil comparé (échelle relative)</div>
          <ResponsiveContainer width="100%" height={340}>
            <RadarChart data={radarData}>
              <PolarGrid stroke={LINE} />
              <PolarAngleAxis dataKey="stat" tick={{ fill: "#8B93A1", fontSize: 11 }} />
              <PolarRadiusAxis tick={{ fill: "#5C6470", fontSize: 9 }} />
              <Radar name={teamA} dataKey={teamA} stroke={AMBER} fill={AMBER} fillOpacity={0.35} />
              <Radar name={teamB} dataKey={teamB} stroke={TEAL} fill={TEAL} fillOpacity={0.25} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Tooltip contentStyle={{ background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, fontSize: 12 }} />
            </RadarChart>
          </ResponsiveContainer>
        </div>
      )}

      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 0.6fr 0.6fr", padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
          <div>Stat</div><div style={{ textAlign: "right" }}>{teamA || "—"}</div><div style={{ textAlign: "right" }}>{teamB || "—"}</div>
        </div>
        {SCOUT_STAT_SCHEMA.map(s => {
          const va = a?.stats[s.key], vb = b?.stats[s.key];
          if (va === undefined && vb === undefined) return null;
          const colorA = statCellColor(s.key, va, vb), colorB = statCellColor(s.key, vb, va);
          const rankA = leagueRank(teams, s.key, teamA), rankB = leagueRank(teams, s.key, teamB);
          return (
            <div key={s.key} style={{ display: "grid", gridTemplateColumns: "1fr 0.6fr 0.6fr", padding: "8px 16px", alignItems: "center", borderBottom: `1px solid ${LINE}`, fontSize: 13 }}>
              <div style={{ color: "#8B93A1" }}>{s.label}</div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, color: colorA || PAPER }}>{va !== undefined ? (s.pct ? va.toFixed(1) + "%" : va.toFixed(1)) : "–"}</div>
                {rankA && <div style={{ fontSize: 9.5, color: "#5C6470" }}>#{rankA.rank}/{rankA.total}</div>}
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, color: colorB || PAPER }}>{vb !== undefined ? (s.pct ? vb.toFixed(1) + "%" : vb.toFixed(1)) : "–"}</div>
                {rankB && <div style={{ fontSize: 9.5, color: "#5C6470" }}>#{rankB.rank}/{rankB.total}</div>}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Observation — importe un fichier de coding pour une équipe adverse observée
// (même format que Import Match) et réutilise directement le même moteur
// d'analyse (fréquence/efficacité/% ouvert) que les fiches joueurs, réparti
// par Plays, Playtypes, Screen defense, Defense type, etc. Les catégories et
// tags reconnus se configurent au même endroit que pour Import Match : Settings.
// ---------------------------------------------------------------------------

function ObservationTab() {
  const [observed, setObserved] = useState({}); // { teamName: { plays, importedAt } }
  const [loading, setLoading] = useState(true);
  const [teamName, setTeamName] = useState("");
  const [preview, setPreview] = useState(null);
  const [fileErr, setFileErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [selected, setSelected] = useState(null);
  const fileRef = useRef();

  useEffect(() => { load(); }, []);
  async function load() {
    setLoading(true);
    setObserved((await storeGet("scouting_observations")) || {});
    setLoading(false);
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileErr(""); setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseMatchFile(buf, currentObservationTagCategories(), "tag");
      if (!parsed.plays.length) throw new Error("No action attributed to a player was found.");
      setPreview({ ...parsed, fileName: file.name });
    } catch (err) { setFileErr(err.message || "Error reading the file."); }
  }

  // BUG RÉEL CORRIGÉ : importer un deuxième fichier pour une équipe déjà observée écrasait
  // complètement le premier import au lieu de s'y ajouter — impossible de combiner plusieurs
  // matchs observés de la même équipe adverse. Chaque import garde maintenant un identifiant,
  // pour pouvoir les combiner (et, si besoin, en retirer un seul plus tard) sans tout perdre.
  async function confirmImport() {
    if (!teamName.trim()) { setFileErr("Enter the observed team's name before confirming."); return; }
    setBusy(true);
    const name = teamName.trim();
    const existing = observed[name] || { plays: [], imports: [] };
    const importId = uid();
    const taggedPlays = preview.plays.map(p => ({ ...p, importId }));
    const next = {
      ...observed,
      [name]: {
        plays: [...(existing.plays || []), ...taggedPlays],
        imports: [...(existing.imports || []), { id: importId, fileName: preview.fileName || "file.xlsx", importedAt: new Date().toISOString(), playsCount: taggedPlays.length }],
        importedAt: new Date().toISOString(), // dernier import, pour l'affichage rétrocompatible
      },
    };
    await storeSet("scouting_observations", next);
    setObserved(next);
    setSelected(name);
    setPreview(null); setTeamName("");
    if (fileRef.current) fileRef.current.value = "";
    setBusy(false);
  }

  // Retire un seul import (un seul fichier) d'une équipe observée, sans toucher aux autres.
  async function removeImport(name, importId) {
    const existing = observed[name];
    if (!existing) return;
    const next = {
      ...observed,
      [name]: {
        ...existing,
        plays: (existing.plays || []).filter(p => p.importId !== importId),
        imports: (existing.imports || []).filter(imp => imp.id !== importId),
      },
    };
    await storeSet("scouting_observations", next);
    setObserved(next);
  }

  async function removeObserved(name) {
    const next = { ...observed };
    delete next[name];
    await storeSet("scouting_observations", next);
    setObserved(next);
    if (selected === name) setSelected(null);
  }

  if (loading) return <EmptyState text="Loading…" />;

  const names = Object.keys(observed);
  const current = selected ? observed[selected] : null;
  const off = current ? current.plays.filter(isOffense) : [];
  const def = current ? current.plays.filter(isDefense) : [];

  return (
    <div>
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 22, marginBottom: 24 }}>
        <p style={{ color: "#8B93A1", fontSize: 13.5, lineHeight: 1.6, margin: "0 0 16px" }}>
          Import a coding file (same format as Import Match) for an opponent team you've scouted. The app breaks
          down their tendencies by frequency and efficiency — plays, playtypes, screen defense, defense type, and
          any other category configured in <b>Settings</b>.
        </p>
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>Observed team name</label>
          <input type="text" list="observed-team-names" placeholder="e.g. Zalgiris U16" value={teamName} onChange={e => setTeamName(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320 }} />
          <datalist id="observed-team-names">
            {Object.keys(observed).map(name => <option key={name} value={name} />)}
          </datalist>
          {teamName.trim() && observed[teamName.trim()] && (
            <div style={{ fontSize: 11.5, color: TEAL, marginTop: 6 }}>
              This team already has {observed[teamName.trim()].imports?.length ?? 1} file(s) imported ({observed[teamName.trim()].plays.length} actions) — this new file will be added to them, not replace them.
            </div>
          )}
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ color: "#8B93A1", fontSize: 13 }} />
        {fileErr && <div style={{ color: RED, fontSize: 13, marginTop: 10 }}>{fileErr}</div>}

        {preview && (
          <div style={{ marginTop: 18, padding: 16, background: PANEL2, borderRadius: 10, border: `1px solid ${LINE}` }}>
            <div style={{ fontSize: 13, color: PAPER, marginBottom: 10 }}>
              Sheet read: <b>{preview.sheetName}</b> · {preview.columnsDetected} columns detected ·
              {" "}<b>{preview.playsWithPlayer}</b> actions recognized / {preview.totalRows} total rows
            </div>
            <button disabled={busy} onClick={confirmImport} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>
              {busy ? "Import…" : "Confirm import"}
            </button>
          </div>
        )}
      </div>

      {names.length > 0 && (
        <div style={{ marginBottom: 24 }}>
          <label style={labelStyle}>Team to view</label>
          <select value={selected || ""} onChange={e => setSelected(e.target.value || null)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 320 }}>
            <option value="">— Choose an observed team —</option>
            {names.map(name => <option key={name} value={name}>{name}</option>)}
          </select>
        </div>
      )}

      {current && (
        <div>
          <SectionTitle eyebrow="Breakdown" title={selected} />
          {current.imports?.length > 0 && (
            <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 10, padding: 14, marginBottom: 20 }}>
              <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#5C6470", marginBottom: 8 }}>
                Files combined for this team ({current.imports.length})
              </div>
              {current.imports.map(imp => (
                <div key={imp.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderTop: `1px solid ${LINE}` }}>
                  <div style={{ fontSize: 12.5, color: "#D8DCE2" }}>
                    {imp.fileName} <span style={{ color: "#5C6470" }}>· {new Date(imp.importedAt).toLocaleDateString()} · {imp.playsCount} actions</span>
                  </div>
                  <button onClick={() => removeImport(selected, imp.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }} title="Remove this file only"><X size={14} /></button>
                </div>
              ))}
            </div>
          )}
          <OffenseDefenseBreakdown off={off} def={def} categories={currentObservationTagCategories()} />
        </div>
      )}

      <div style={{ height: 8 }} />
      <SectionTitle eyebrow="History" title="Observed teams" />
      {names.length === 0 ? <EmptyState text="No team observed yet." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {names.map(name => (
            <div key={name} style={{ ...btnRow, cursor: "default" }}>
              <span>{name} <span style={{ color: "#5C6470", fontSize: 12 }}>· {observed[name].plays.length} actions · imported {new Date(observed[name].importedAt).toLocaleDateString("en-US")}</span></span>
              <button onClick={() => removeObserved(name)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={15} /></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ScoutingTeamRow({ name, team, onSaveLogo, onDelete }) {
  const logoRef = useRef();
  async function handleLogo(e) {
    const file = e.target.files[0];
    if (!file) return;
    onSaveLogo(await fileToResizedDataURL(file, 200, 0.9));
  }
  return (
    <div style={{ ...btnRow, cursor: "default" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <button type="button" onClick={() => logoRef.current && logoRef.current.click()} title="Set logo" style={{ width: 26, height: 26, borderRadius: 6, background: PANEL2, border: `1px solid ${LINE}`, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0, cursor: "pointer", padding: 0 }}>
          {team.logo ? <img src={team.logo} alt="" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> : <Camera size={12} color="#5C6470" />}
        </button>
        <input ref={logoRef} type="file" accept="image/*" onChange={handleLogo} style={{ display: "none" }} />
        <span>{name} <span style={{ color: "#5C6470", fontSize: 12 }}>· source: {team.source} · updated {team.updatedAt}</span></span>
      </div>
      <button onClick={onDelete} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={15} /></button>
    </div>
  );
}

function ScoutingTab({ isCoach, matchFilter, initialSubtab, initialReportTeam }) {
  const scouting = useScoutingTeams();
  const advanced = useTeamAdvancedStats(matchFilter);
  const box = useAllBoxScores(matchFilter);
  const [subtab, setSubtab] = useState(initialSubtab || "comparaison");
  const [showManual, setShowManual] = useState(false);
  const [pending, setPending] = useState(null); // { teams: [...], kind } en attente de décision remplacer/ajouter
  const [replaceChoice, setReplaceChoice] = useState(null); // "replace" | "merge"
  const [busy, setBusy] = useState(false);
  const [fileErr, setFileErr] = useState("");
  const fileOffRef = useRef();
  const fileDefRef = useRef();

  // "Our team" est injectée automatiquement à partir des box scores déjà importés,
  // en plus des équipes de scouting enregistrées manuellement/via fichier/photo.
  const ourStats = ourTeamAsScoutStats(advanced, box);
  const allTeams = { ...(Object.keys(ourStats).length ? { "Our team": { stats: ourStats, source: "calculated" } } : {}), ...scouting.teams };

  function startPending(teams, kind) {
    setPending({ teams: teams.map(t => ({ ...t, selected: true })), kind });
    setReplaceChoice(null);
  }

  async function handleExcelFile(e, side) {
    const file = e.target.files[0];
    if (!file) return;
    setFileErr("");
    try {
      const buf = await file.arrayBuffer();
      // Deux formats reconnus automatiquement : "Stats Center" (colonne Club) et
      // leaderboard FIBA officiel (colonne Team, en anglais) — on essaie l'un puis l'autre.
      let parsed;
      try { parsed = parseScoutingExcelFile(buf); }
      catch { parsed = parseFibaLeaderboardFile(buf, side); }
      if (!parsed.teams.length) throw new Error("No team detected in this file.");
      startPending(parsed.teams, side === "defense" ? "excel-def" : "excel-off");
    } catch (err) { setFileErr(err.message || "File format not recognized — tell me how it's structured and I'll add detection for it."); }
    e.target.value = "";
  }

  async function confirmPending() {
    setBusy(true);
    if (replaceChoice === "replace") {
      for (const name of Object.keys(scouting.teams)) await scouting.deleteTeam(name);
    }
    for (const t of pending.teams.filter(x => x.selected)) {
      // Import défensif : on FUSIONNE avec les stats offensives déjà enregistrées pour cette
      // équipe (même nom) plutôt que de les écraser, pour obtenir une fiche complète.
      const existing = scouting.teams[t.name];
      const mergedStats = existing && pending.kind === "excel-def" ? { ...existing.stats, ...t.stats } : t.stats;
      await scouting.saveTeam(t.name, mergedStats, pending.kind.startsWith("excel") ? "excel" : "photo");
    }
    setBusy(false); setPending(null); setReplaceChoice(null);
  }

  return (
    <div>
      <SectionTitle eyebrow="Scouting" title="Team comparison" />
      <div style={{ display: "flex", gap: 8, marginBottom: 20, borderBottom: `1px solid ${LINE}`, paddingBottom: 10, flexWrap: "wrap" }}>
        {[["comparaison", "Comparison"], ["rapport", "Scouting Report"], ...(isCoach ? [["observation", "Observation"]] : []), ...(isCoach ? [["gerer", "Manage teams"]] : [])].map(([id, label]) => (
          <button key={id} onClick={() => setSubtab(id)} style={{
            padding: "7px 12px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "inherit",
            background: subtab === id ? PANEL2 : "transparent", color: subtab === id ? AMBER : "#8B93A1"
          }}>{label}</button>
        ))}
      </div>

      {subtab === "comparaison" && (scouting.loading ? <EmptyState text="Loading…" /> : <ScoutingComparison teams={allTeams} />)}
      {subtab === "rapport" && <ScoutingReportTab isCoach={isCoach} teamNames={Object.keys(scouting.teams)} scoutingTeams={scouting.teams} onSaveLogo={scouting.saveLogo} initialTeam={initialReportTeam} />}
      {subtab === "observation" && isCoach && <ObservationTab />}

      {subtab === "gerer" && isCoach && (
        <div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
            <button onClick={() => fileOffRef.current && fileOffRef.current.click()} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
              <Upload size={14} /> Import file — Offense
            </button>
            <input ref={fileOffRef} type="file" accept=".xlsx,.xls" onChange={e => handleExcelFile(e, "offense")} style={{ display: "none" }} />
            <button onClick={() => fileDefRef.current && fileDefRef.current.click()} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
              <Upload size={14} /> Import file — Defense
            </button>
            <input ref={fileDefRef} type="file" accept=".xlsx,.xls" onChange={e => handleExcelFile(e, "defense")} style={{ display: "none" }} />
            <button onClick={() => setShowManual(s => !s)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
              <Plus size={14} /> Manual entry
            </button>
          </div>
          <div style={{ fontSize: 11.5, color: "#5C6470", marginBottom: 14 }}>
            "Offense" file: Poss, Pts, %2pt, %3pt, ORTG… "Defense" file (same FIBA format, stats allowed):
            DRTG is automatically derived from its PPP column. Import both for the same team to complete its profile.
          </div>
          {fileErr && <div style={{ color: RED, fontSize: 13, marginBottom: 14 }}>{fileErr}</div>}

          <ScoutingPhotoImport onExtracted={(teams) => startPending(teams, "photo")} />

          {pending && replaceChoice === null && (
            <div style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: 16, background: PANEL2, border: `1px solid ${AMBER}`, borderRadius: 10, marginBottom: 20 }}>
              <AlertTriangle size={16} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 6 }}>{pending.teams.length} team(s) found in this file.</div>
                <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 12 }}>Do you want to delete all previously saved teams before importing, or keep the existing ones and only add/update?</div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => setReplaceChoice("replace")} style={{ ...btnPrimary, width: "auto", padding: "9px 16px", background: RED, color: "#fff" }}>Yes, replace everything</button>
                  <button onClick={() => setReplaceChoice("merge")} style={{ ...btnSecondary }}>No, add / update</button>
                  <button onClick={() => setPending(null)} style={{ ...btnSecondary }}>Cancel import</button>
                </div>
              </div>
            </div>
          )}

          {pending && replaceChoice !== null && (
            <div style={{ background: PANEL2, border: `1px solid ${TEAL}`, borderRadius: 10, padding: 16, marginBottom: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>
                {replaceChoice === "replace" ? "All existing teams will be deleted, then:" : "Add / update:"}
              </div>
              <div style={{ maxHeight: 260, overflowY: "auto", marginBottom: 10 }}>
                {pending.teams.map((t, i) => (
                  <label key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, padding: "4px 0" }}>
                    <input type="checkbox" checked={t.selected} onChange={e => setPending(p => ({ ...p, teams: p.teams.map((x, j) => j === i ? { ...x, selected: e.target.checked } : x) }))} />
                    {t.name}
                  </label>
                ))}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => { setPending(null); setReplaceChoice(null); }} style={btnSecondary}>Cancel</button>
                <button disabled={busy} onClick={confirmPending} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>{busy ? "…" : `Save (${pending.teams.filter(t => t.selected).length})`}</button>
              </div>
            </div>
          )}

          {showManual && (
            <ScoutingTeamForm
              busy={busy}
              onCancel={() => setShowManual(false)}
              onSave={async (name, stats) => { setBusy(true); await scouting.saveTeam(name, stats, "manual"); setBusy(false); setShowManual(false); }}
            />
          )}

          <SectionTitle eyebrow="Saved teams" title={`${Object.keys(scouting.teams).length} team(s)`} />
          {Object.keys(scouting.teams).length === 0 ? <EmptyState text="No scouting team saved yet." /> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {Object.entries(scouting.teams).map(([name, t]) => (
                <ScoutingTeamRow key={name} name={name} team={t} onSaveLogo={logo => scouting.saveLogo(name, logo)} onDelete={() => scouting.deleteTeam(name)} />
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Sauvegarde — export/import complet des données de l'équipe active
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Settings — catégories de colonnes du fichier de coding (playtypes, plays, joueur,
// erreurs défensives, sélection de tir…), modifiables sans toucher au code.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Settings — box score column mapping (which exact column names count as which stat)
// ---------------------------------------------------------------------------

function BoxColumnAliasesSettings() {
  const [aliases, setAliases] = useState(boxColumnAliases());
  const [newAliasByKey, setNewAliasByKey] = useState({});
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");

  async function persist(next) {
    setBusy(true);
    await saveBoxColumnAliases(next);
    setAliases(next);
    setStatus("Saved — applied immediately, including for future imports.");
    setBusy(false);
  }

  function addAlias(key) {
    const val = (newAliasByKey[key] || "").trim();
    if (!val) return;
    const next = { ...aliases, [key]: [...(aliases[key] || []), val] };
    setNewAliasByKey(s => ({ ...s, [key]: "" }));
    persist(next);
  }
  function removeAlias(key, alias) {
    persist({ ...aliases, [key]: (aliases[key] || []).filter(a => a !== alias) });
  }

  const keys = Object.keys(STAT_KEY_FRIENDLY_NAME);

  return (
    <div>
      <SectionTitle eyebrow="Settings" title="Box score column mapping" />
      <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, marginBottom: 20 }}>
        The app already recognizes many common column names automatically. If a box score file uses a column name
        it doesn't recognize (for example a stat showing up as "–" that should have a real value), add the exact
        column name here, under the matching stat — it will then be recognized with certainty on every future import.
      </p>

      {keys.map(key => {
        const list = aliases[key] || [];
        return (
          <div key={key} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 14 }}>
            <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>{STAT_KEY_FRIENDLY_NAME[key]}</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
              {list.length === 0 && <span style={{ fontSize: 12.5, color: "#5C6470" }}>No custom column name added yet — automatic detection is used.</span>}
              {list.map(alias => (
                <span key={alias} style={{ display: "flex", alignItems: "center", gap: 6, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 7, padding: "5px 10px", fontSize: 12.5 }}>
                  {alias}
                  <button onClick={() => removeAlias(key, alias)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex", padding: 0 }}><X size={12} /></button>
                </span>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <input value={newAliasByKey[key] || ""} onChange={e => setNewAliasByKey(s => ({ ...s, [key]: e.target.value }))}
                onKeyDown={e => { if (e.key === "Enter") addAlias(key); }}
                placeholder="Exact column name in your file" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", flex: 1, maxWidth: 280 }} />
              <button disabled={busy} onClick={() => addAlias(key)} style={{ ...btnSecondary, display: "flex", alignItems: "center", gap: 6 }}><Plus size={13} /> Add</button>
            </div>
          </div>
        );
      })}
      {status && <div style={{ fontSize: 12.5, color: TEAL }}>{status}</div>}
    </div>
  );
}

function TagCategoriesSettings({ roster, title = "Column categories (coding file)", getCurrent = currentTagCategories, onSave = saveTagCategories, onResetDefault = () => JSON.parse(JSON.stringify(DEFAULT_TAG_CATEGORIES)) }) {
  const [cats, setCats] = useState(getCurrent());
  const [newTagByCat, setNewTagByCat] = useState({});
  const [newCatName, setNewCatName] = useState("");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [chartStyles, setChartStyles] = useState(currentCategoryChartStyles());

  const BUILTIN_CATEGORIES = new Set(["Player", "Playtypes", "Plays", "Shot selection", "Defensive mistakes", "Screen defense", "Spacing", "Shot zone", "Results & misc."]);

  async function setChartStyle(catName, style) {
    const next = { ...chartStyles, [catName]: style };
    setChartStyles(next);
    await saveCategoryChartStyles(next);
  }

  async function persist(next) {
    setBusy(true);
    await onSave(next);
    setCats(next);
    setStatus("Saved — applied immediately, including for future imports.");
    setBusy(false);
  }

  function syncPlayersFromRoster() {
    const names = Array.from(new Set([...(cats["Player"] || []), ...roster.map(p => p.name)]));
    persist({ ...cats, "Player": names });
  }

  function addTag(catName) {
    const val = (newTagByCat[catName] || "").trim();
    if (!val) return;
    const next = { ...cats, [catName]: [...cats[catName], val] };
    setNewTagByCat(s => ({ ...s, [catName]: "" }));
    persist(next);
  }
  function removeTag(catName, tag) {
    persist({ ...cats, [catName]: cats[catName].filter(t => t !== tag) });
  }
  function addCategory() {
    const name = newCatName.trim();
    if (!name || cats[name]) return;
    persist({ ...cats, [name]: [] });
    setNewCatName("");
  }
  function removeCategory(name) {
    const next = { ...cats };
    delete next[name];
    persist(next);
  }
  function resetDefaults() {
    persist(onResetDefault());
  }

  return (
    <div>
      <SectionTitle eyebrow="Settings" title={title} />
      <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, marginBottom: 20 }}>
        Every column in your coding file must belong to a category to be read correctly — otherwise it risks
        being mistaken for a player. If your coding software changes (new column, renamed one), add it
        here to the right category, or create a new category if needed — it will then also appear in the
        Offense/Defense charts on player and team pages. The starting values reproduce exactly what was
        already in use — nothing changes unless you edit something here.
      </p>

      <div style={{ display: "flex", gap: 10, alignItems: "flex-end", marginBottom: 20, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 220 }}>
          <label style={labelStyle}>New category</label>
          <input value={newCatName} onChange={e => setNewCatName(e.target.value)} placeholder="e.g. Defensive transition" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
        </div>
        <button disabled={busy || !newCatName.trim()} onClick={addCategory} style={{ ...btnPrimary, width: "auto", padding: "10px 18px", display: "flex", alignItems: "center", gap: 6 }}>
          <Plus size={14} /> Add category
        </button>
        <button onClick={resetDefaults} style={btnSecondary}>Reset to default values</button>
      </div>
      {status && <div style={{ fontSize: 12.5, color: TEAL, marginBottom: 16 }}>{status}</div>}

      {Object.entries(cats).map(([catName, tags]) => (
        <div key={catName} style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 14 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontSize: 14, fontWeight: 700 }}>{catName}{catName === "Player" && <span style={{ fontSize: 11, color: "#5C6470", fontWeight: 400 }}> — first names recognized with certainty as players</span>}</div>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              {catName === "Player" && (
                <button onClick={syncPlayersFromRoster} style={{ fontSize: 11.5, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>Sync with roster</button>
              )}
              {catName !== "Player" && (
                <button onClick={() => removeCategory(catName)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }} title="Delete category"><Trash2 size={14} /></button>
              )}
            </div>
          </div>
          {!BUILTIN_CATEGORIES.has(catName) && (
            <div style={{ marginBottom: 12 }}>
              <label style={labelStyle}>Chart shown for this category</label>
              <select value={chartStyles[catName] || "detailed"} onChange={e => setChartStyle(catName, e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 340 }}>
                <option value="detailed">Detailed list — frequency, PPP and % open per tag (like Plays)</option>
                <option value="simple">Simple comparison — a donut comparing the tags to each other (like Shooting Selection)</option>
                <option value="both">Both — show the donut and the detailed list together</option>
              </select>
            </div>
          )}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
            {tags.length === 0 && <span style={{ fontSize: 12.5, color: "#5C6470" }}>No column yet.</span>}
            {tags.map(tag => (
              <span key={tag} style={{ display: "flex", alignItems: "center", gap: 6, background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 7, padding: "5px 10px", fontSize: 12.5 }}>
                {tag}
                <button onClick={() => removeTag(catName, tag)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex", padding: 0 }}><X size={12} /></button>
              </span>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input value={newTagByCat[catName] || ""} onChange={e => setNewTagByCat(s => ({ ...s, [catName]: e.target.value }))}
              onKeyDown={e => { if (e.key === "Enter") addTag(catName); }}
              placeholder="Exact column name (e.g. 3PT+)" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", flex: 1, maxWidth: 280 }} />
            <button disabled={busy} onClick={() => addTag(catName)} style={{ ...btnSecondary, display: "flex", alignItems: "center", gap: 6 }}><Plus size={13} /> Add</button>
          </div>
        </div>
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Migration prénom → nom complet — les données déjà enregistrées avant le passage à
// l'identité par nom complet restent accessibles sous l'ancienne clé (prénom seul) tant
// qu'on ne les recopie pas explicitement. Cette fonction copie tout ce qui existe déjà
// (entraînement, objectifs, Wellness, évaluation mentale, photo, rôle, identifiants de
// connexion) vers la nouvelle clé, et corrige aussi les joueurs déjà référencés à
// l'intérieur des matchs codés et box scores déjà importés.
// ---------------------------------------------------------------------------

async function migrateToFullNameIdentity(roster) {
  let migrated = 0;
  const perPlayerKeys = ["training", "objectives", "mental", "wellness", "photo", "role", "meetings"];
  for (const p of roster) {
    if (!p.first || !p.name || p.first === p.name) continue;
    for (const prefix of perPlayerKeys) {
      const oldData = await storeGet(prefix + ":" + p.first);
      const newData = await storeGet(prefix + ":" + p.name);
      if (oldData && !newData) { await storeSet(prefix + ":" + p.name, oldData); migrated++; }
    }
  }
  // Identifiants de connexion (même code PIN conservé, juste rattaché au nom complet).
  const users = (await storeGet("app_users")) || {};
  let usersChanged = false;
  for (const p of roster) {
    if (p.first && p.name && p.first !== p.name && users[p.first] && !users[p.name]) {
      users[p.name] = users[p.first];
      usersChanged = true;
      migrated++;
    }
  }
  if (usersChanged) await storeSet("app_users", users);

  // Matchs déjà codés : chaque action porte le prénom tel qu'il apparaissait dans l'ancien
  // fichier — on le remplace par le nom complet du joueur correspondant du roster actuel.
  const matchIdx = (await storeGet("match_index")) || [];
  for (const m of matchIdx) {
    const data = await storeGet("match:" + m.id);
    if (!data || !data.plays) continue;
    let changed = false;
    const plays = data.plays.map(play => {
      const p = roster.find(pl => pl.first === play.player && pl.first !== pl.name);
      if (p) { changed = true; return { ...play, player: p.name }; }
      return play;
    });
    if (changed) { await storeSet("match:" + m.id, { ...data, plays }); migrated++; }
  }

  // Box scores déjà importés : idem pour chaque ligne joueur.
  const boxIdx = (await storeGet("boxscore_index")) || [];
  for (const b of boxIdx) {
    const data = await storeGet("boxscore:" + b.id);
    if (!data || !data.rows) continue;
    let changed = false;
    const rows = data.rows.map(row => {
      const p = roster.find(pl => pl.first === row.player && pl.first !== pl.name);
      if (p) { changed = true; return { ...row, player: p.name, playerFull: p.name }; }
      return row;
    });
    if (changed) { await storeSet("boxscore:" + b.id, { ...data, rows }); migrated++; }
  }

  return migrated;
}

// Migration lors du RENOMMAGE d'un joueur (pas seulement prénom → nom complet, mais n'importe
// quel changement de nom) : reprend tout ce qui était enregistré sous l'ancien nom et le
// rattache au nouveau, pour qu'aucune donnée déjà rentrée ne devienne orpheline.
async function migratePlayerRename(oldName, newName) {
  if (!oldName || !newName || oldName === newName) return;
  const perPlayerKeys = ["training", "objectives", "mental", "wellness", "photo", "role", "meetings"];
  for (const prefix of perPlayerKeys) {
    const oldData = await storeGet(prefix + ":" + oldName);
    if (oldData !== null && oldData !== undefined) {
      await storeSet(prefix + ":" + newName, oldData);
      await storeDelete(prefix + ":" + oldName);
    }
  }
  // Identifiants de connexion (même code PIN conservé, juste rattaché au nouveau nom).
  const users = (await storeGet("app_users")) || {};
  if (users[oldName]) {
    users[newName] = users[oldName];
    delete users[oldName];
    await storeSet("app_users", users);
  }
  // Matchs déjà codés : remplace l'ancien nom par le nouveau dans chaque action du joueur.
  const matchIdx = (await storeGet("match_index")) || [];
  for (const m of matchIdx) {
    const data = await storeGet("match:" + m.id);
    if (!data || !data.plays) continue;
    if (!data.plays.some(play => play.player === oldName)) continue;
    const plays = data.plays.map(play => play.player === oldName ? { ...play, player: newName } : play);
    await storeSet("match:" + m.id, { ...data, plays });
  }
  // Box scores déjà importés : idem pour chaque ligne joueur.
  const boxIdx = (await storeGet("boxscore_index")) || [];
  for (const b of boxIdx) {
    const data = await storeGet("boxscore:" + b.id);
    if (!data || !data.rows) continue;
    if (!data.rows.some(row => row.player === oldName)) continue;
    const rows = data.rows.map(row => row.player === oldName ? { ...row, player: newName, playerFull: newName } : row);
    await storeSet("boxscore:" + b.id, { ...data, rows });
  }
}

function BackupTab({ team, roster }) {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [pendingImport, setPendingImport] = useState(null); // backup lu, en attente de confirmation
  const [error, setError] = useState("");
  const [backupText, setBackupText] = useState(""); // solution de secours : texte copiable si le téléchargement ne fonctionne pas
  const [migrateBusy, setMigrateBusy] = useState(false);
  const [migrateStatus, setMigrateStatus] = useState("");
  const fileRef = useRef();
  const textRef = useRef();

  async function runMigration() {
    setMigrateBusy(true); setMigrateStatus("");
    const count = await migrateToFullNameIdentity(roster);
    setMigrateStatus(count > 0
      ? `Done — ${count} item(s) migrated to the full-name identity.`
      : "Nothing to migrate — everything is already using the full name.");
    setMigrateBusy(false);
  }
  const pasteRef = useRef();

  async function handleExport() {
    setBusy(true); setStatus(""); setError("");
    try {
      const backup = await exportAllData(team.id, team.name);
      const json = JSON.stringify(backup, null, 2);
      setBackupText(json); // toujours prêt, même si le téléchargement échoue silencieusement
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `backup_${team.name.replace(/[^a-z0-9]+/gi, "_")}_${todayLocal()}.json`;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      setStatus(`Backup ready — ${Object.keys(backup.data).length} items. If the download didn't start, use "Copy text" below.`);
    } catch (err) { setError("Export failed: " + (err.message || "unknown error")); }
    setBusy(false);
  }

  function copyBackupText() {
    if (textRef.current) {
      textRef.current.select();
      try { document.execCommand("copy"); setStatus("Text copied — paste it into a text file and save it as .json."); }
      catch { setStatus("Select the text below and copy it manually (Ctrl+C / Cmd+C)."); }
    }
  }

  async function handleFilePick(e) {
    const file = e.target.files[0];
    if (!file) return;
    setError(""); setStatus("");
    try {
      const text = await file.text();
      loadBackupText(text);
    } catch (err) { setError("Unreadable file: " + (err.message || "invalid format")); }
    e.target.value = "";
  }

  function loadBackupText(text) {
    const backup = JSON.parse(text);
    if (!backup || typeof backup.data !== "object") throw new Error("This text doesn't look like a valid backup of the app.");
    setPendingImport(backup);
  }

  function handlePasteImport() {
    setError(""); setStatus("");
    const text = (pasteRef.current && pasteRef.current.value || "").trim();
    if (!text) { setError("Paste the backup content into the text box first."); return; }
    try { loadBackupText(text); } catch (err) { setError("Unreadable text: " + (err.message || "invalid format")); }
  }

  async function confirmImport() {
    setBusy(true);
    try {
      const count = await importAllData(team.id, pendingImport);
      setStatus(`Import complete — ${count} items restored. Reload the page (or switch and come back to this team) to see everything up to date.`);
    } catch (err) { setError("Import failed: " + (err.message || "unknown error")); }
    setBusy(false); setPendingImport(null);
  }

  return (
    <div>
      <SectionTitle eyebrow="Data protection" title="Backup" />

      <div style={{ background: PANEL, border: `1px solid ${AMBER}`, borderRadius: 12, padding: 22, marginBottom: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>Migrate to full-name identity</div>
        <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, margin: "0 0 16px" }}>
          Players are now identified by their full name everywhere (to avoid mixing up two players
          who share a first name). If you entered training, objectives, Wellness, mental evaluations,
          or imported matches/box scores before this change, run this once to carry everything over —
          nothing already there gets lost.
        </p>
        <button disabled={migrateBusy} onClick={runMigration} style={{ ...btnPrimary, width: "auto", padding: "10px 20px" }}>
          {migrateBusy ? "Migrating…" : "Run migration"}
        </button>
        {migrateStatus && <div style={{ fontSize: 13, color: TEAL, marginTop: 10 }}>{migrateStatus}</div>}
      </div>

      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 22, marginBottom: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>Export</div>
        <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, margin: "0 0 16px" }}>
          Downloads a file containing absolutely everything for <b>{team.name}</b>: roster, coded matches, box scores,
          scouting (comparisons + reports), objectives, training, mental evaluations, photos. Do this regularly,
          and keep it somewhere outside the app (computer, cloud) — it's your only safeguard in case of a problem
          with the app's storage.
        </p>
        <button disabled={busy} onClick={handleExport} style={{ ...btnPrimary, width: "auto", padding: "10px 20px", display: "flex", alignItems: "center", gap: 8 }}>
          <Download size={15} /> {busy ? "…" : "Download backup"}
        </button>

        {backupText && (
          <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${LINE}` }}>
            <div style={{ display: "flex", gap: 8, alignItems: "flex-start", padding: 10, background: PANEL2, border: `1px solid ${AMBER}`, borderRadius: 8, marginBottom: 10 }}>
              <AlertTriangle size={14} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
              <div style={{ fontSize: 12, color: "#D8DCE2", lineHeight: 1.5 }}>
                If no file downloaded automatically (this can happen depending on the environment), here's a fallback
                that doesn't depend on downloading: click the text below, select all (Ctrl/Cmd+A),
                copy (Ctrl/Cmd+C), paste into a text editor, and save the file with the extension <b>.json</b>.
              </div>
            </div>
            <button onClick={copyBackupText} style={{ ...btnSecondary, marginBottom: 8 }}>Copy text automatically</button>
            <textarea ref={textRef} readOnly value={backupText} onClick={e => e.target.select()}
              style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "ui-monospace, monospace", fontSize: 11, height: 160, resize: "vertical", whiteSpace: "pre" }} />
          </div>
        )}
      </div>

      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 22, marginBottom: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>Restore</div>
        <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, margin: "0 0 16px" }}>
          Reload a previously downloaded backup file. <b>Overwrites existing data</b> with the same
          name (same matches, same roster...) — only imports into the currently selected team ({team.name}).
        </p>
        <input ref={fileRef} type="file" accept=".json" onChange={handleFilePick} style={{ color: "#8B93A1", fontSize: 13 }} />
        <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${LINE}` }}>
          <div style={{ fontSize: 12, color: "#8B93A1", marginBottom: 8 }}>— OR, if the file never downloaded — paste the text you copied here:</div>
          <textarea ref={pasteRef} placeholder="Paste the backup's JSON content here…" rows={4}
            style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "ui-monospace, monospace", fontSize: 11, resize: "vertical", marginBottom: 8 }} />
          <button onClick={handlePasteImport} style={btnSecondary}>Read this text</button>
        </div>
      </div>

      {pendingImport && (
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: 16, background: PANEL2, border: `1px solid ${AMBER}`, borderRadius: 10, marginBottom: 20 }}>
          <AlertTriangle size={16} color={AMBER} style={{ marginTop: 2, flexShrink: 0 }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 6 }}>
              Backup from {pendingImport.exportedAt ? new Date(pendingImport.exportedAt).toLocaleDateString("en-US") : "?"}
              {pendingImport.teamName ? ` (original team: ${pendingImport.teamName})` : ""} — {Object.keys(pendingImport.data).length} items.
            </div>
            <div style={{ fontSize: 12.5, color: "#8B93A1", marginBottom: 12 }}>
              Confirm to restore this data into <b>{team.name}</b> (the team currently open). Existing data
              with the same keys will be replaced.
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => setPendingImport(null)} style={btnSecondary}>Cancel</button>
              <button disabled={busy} onClick={confirmImport} style={{ ...btnPrimary, width: "auto", padding: "9px 16px", background: RED, color: "#fff" }}>{busy ? "…" : "Yes, restore and overwrite"}</button>
            </div>
          </div>
        </div>
      )}

      {status && <div style={{ fontSize: 13, color: TEAL, marginBottom: 12 }}>{status}</div>}
      {error && <div style={{ fontSize: 13, color: RED, marginBottom: 12 }}>{error}</div>}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Training collectif — une séance appliquée à plusieurs joueurs d'un coup,
// avec les mêmes infos partagées mais une note individuelle par joueur.
// ---------------------------------------------------------------------------

function CollectiveTraining({ roster }) {
  const [selected, setSelected] = useState([]); // ids de joueurs
  const [date, setDate] = useState(todayLocal());
  const { themes, addTheme } = useTrainingThemes();
  const [thematique, setThematique] = useState(themes[0]);
  const [newThemeName, setNewThemeName] = useState("");
  const [theme, setTheme] = useState("");
  const [objectif, setObjective] = useState("");
  const [duree, setDuree] = useState(15);
  const [commentaire, setCommentaire] = useState("");
  const [notes, setNotes] = useState({}); // playerId -> note 1-5
  const [comments, setComments] = useState({}); // playerId -> commentaire individuel
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  function toggle(id) {
    setSelected(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
    setNotes(n => (id in n ? n : { ...n, [id]: null }));
    setComments(c => (id in c ? c : { ...c, [id]: "" }));
  }
  function selectAll() {
    setSelected(roster.map(p => p.id));
    setNotes(Object.fromEntries(roster.map(p => [p.id, null])));
    setComments(c => Object.fromEntries(roster.map(p => [p.id, c[p.id] || ""])));
  }
  function selectNone() { setSelected([]); }

  async function submit() {
    setError("");
    if (!selected.length) { setError("Select at least one player."); return; }
    if (!theme.trim()) { setError("Enter a specific theme before confirming."); return; }
    setBusy(true); setStatus("");
    try {
      for (const id of selected) {
        const player = roster.find(p => p.id === id);
        if (!player) continue;
        const existing = (await storeGet("training:" + player.name)) || [];
        const individualComment = (comments[id] || "").trim();
        const fullComment = [commentaire, individualComment].filter(Boolean).join(commentaire && individualComment ? "\n\n" : "");
        const entry = { id: uid(), date, thematique, theme, objectif, duree, commentaire: fullComment, eval: notes[id] ?? null };
        await storeSet("training:" + player.name, [entry, ...existing]);
      }
      setStatus(`Session added to ${selected.length} player${selected.length !== 1 ? "s" : ""} — saved under: ${selected.map(id => { const p = roster.find(pl => pl.id === id); return p ? `"${p.name}"` : "?"; }).join(", ")}.`);
      setTheme(""); setObjective(""); setCommentaire(""); setComments({});
    } catch (err) {
      setError("Save failed: " + (err.message || "erreur inconnue"));
    }
    setBusy(false);
  }

  return (
    <div>
      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Shared session info</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
          <div><label style={labelStyle}>Date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 160 }} /></div>
          <div style={{ width: 180 }}>
            <label style={labelStyle}>Category</label>
            <select value={thematique} onChange={e => setThematique(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", color: trainingThemeColor(thematique, themes), fontWeight: 700 }}>
              {themes.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
            <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
              <input value={newThemeName} onChange={e => setNewThemeName(e.target.value)} placeholder="New category…" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", fontSize: 11.5, padding: "6px 8px" }} />
              <button type="button" onClick={async () => { await addTheme(newThemeName); setNewThemeName(""); }} style={{ ...btnSecondary, padding: "6px 10px", fontSize: 11.5 }}>Add</button>
            </div>
          </div>
          <div style={{ flex: 1, minWidth: 160 }}><label style={labelStyle}>Specific theme</label><input value={theme} onChange={e => setTheme(e.target.value)} placeholder="e.g. Screen defense" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
          <div style={{ width: 110 }}><label style={labelStyle}>Duration (min)</label><input type="number" min={0} value={duree} onChange={e => setDuree(Number(e.target.value))} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
        </div>
        <div style={{ marginBottom: 10 }}><label style={labelStyle}>Objective</label><input value={objectif} onChange={e => setObjective(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} /></div>
        <div><label style={labelStyle}>Comments (shared)</label><textarea value={commentaire} onChange={e => setCommentaire(e.target.value)} rows={2} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical" }} /></div>
      </div>

      <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontSize: 13, fontWeight: 700 }}>Players involved — individual rating & comment per player</div>
          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={selectAll} style={{ fontSize: 11.5, color: AMBER, background: "none", border: "none", cursor: "pointer" }}>All</button>
            <button onClick={selectNone} style={{ fontSize: 11.5, color: "#5C6470", background: "none", border: "none", cursor: "pointer" }}>None</button>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {roster.map(p => (
            <div key={p.id} style={{ padding: "8px 12px", background: selected.includes(p.id) ? PANEL2 : "transparent", border: `1px solid ${LINE}`, borderRadius: 8 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 10, flex: 1, cursor: "pointer" }}>
                  <input type="checkbox" checked={selected.includes(p.id)} onChange={() => toggle(p.id)} />
                  <span style={{ fontSize: 13.5 }}>{p.name}</span>
                </label>
                {selected.includes(p.id) && (
                  <select value={notes[p.id] ?? ""} onChange={e => setNotes(n => ({ ...n, [p.id]: e.target.value === "" ? null : Number(e.target.value) }))}
                    style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: 100, padding: "6px 10px", fontSize: 13, color: ratingColor(notes[p.id] ?? null), fontWeight: 700 }}>
                    <option value="">No rating</option>
                    {[1, 2, 3, 4, 5].map(v => <option key={v} value={v}>Rating {v}</option>)}
                  </select>
                )}
              </div>
              {selected.includes(p.id) && (
                <textarea
                  value={comments[p.id] || ""}
                  onChange={e => setComments(c => ({ ...c, [p.id]: e.target.value }))}
                  placeholder={`Individual comment for ${p.first}…`}
                  rows={2}
                  style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", resize: "vertical", marginTop: 8, fontSize: 12.5 }}
                />
              )}
            </div>
          ))}
        </div>
      </div>

      <button disabled={busy} onClick={submit} style={{ ...btnPrimary, width: "auto", padding: "11px 22px", display: "flex", alignItems: "center", gap: 8, opacity: busy ? 0.6 : 1 }}>
        <Plus size={15} /> {busy ? "…" : `Add the session to ${selected.length} player${selected.length !== 1 ? "s" : ""}`}
      </button>
      {status && <div style={{ fontSize: 13, color: TEAL, marginTop: 10 }}>{status}</div>}
      {error && <div style={{ fontSize: 13, color: RED, marginTop: 10 }}>{error}</div>}
    </div>
  );
}

function TeamPrintReport({ team, rows, otherLabels, advanced, teamOff, teamDef, roster }) {
  const [resources, setResources] = useState([]);

  useEffect(() => {
    storeGet("team_resources").then(r => setResources(((r || []).sort((a, b) => b.addedAt.localeCompare(a.addedAt)))));
  }, []);

  return (
    <div style={{ padding: 24, background: "#ffffff", color: "#1A1D24" }}>
      <h1 style={{ fontSize: 22, marginBottom: 4 }}>{team?.name || "Team"}</h1>
      <div style={{ fontSize: 12, color: "#8B93A1", marginBottom: 20 }}>Report generated on {new Date().toLocaleDateString("en-US")}</div>

      <h2 style={{ fontSize: 16, marginBottom: 8 }}>Standings</h2>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 10, marginBottom: 20, tableLayout: "fixed" }}>
        <thead><tr>
          <th style={{ border: `1px solid ${LINE}`, padding: "3px 4px", textAlign: "left" }}>#</th>
          <th style={{ border: `1px solid ${LINE}`, padding: "3px 4px", textAlign: "left" }}>Player</th>
          <th style={{ border: `1px solid ${LINE}`, padding: "3px 4px", textAlign: "left" }}>Games</th>
          <th style={{ border: `1px solid ${LINE}`, padding: "3px 4px", textAlign: "left" }}>Pts/game</th>
          {otherLabels.map(l => <th key={l} style={{ border: `1px solid ${LINE}`, padding: "3px 4px", textAlign: "left" }}>{l}</th>)}
        </tr></thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.name}>
              <td style={{ border: `1px solid ${LINE}`, padding: "3px 4px" }}>{i + 1}</td>
              <td style={{ border: `1px solid ${LINE}`, padding: "3px 4px" }}>{r.name}</td>
              <td style={{ border: `1px solid ${LINE}`, padding: "3px 4px" }}>{r.games}</td>
              <td style={{ border: `1px solid ${LINE}`, padding: "3px 4px" }}>{r.ppg !== null ? r.ppg.toFixed(1) : "–"}</td>
              {r.others.map((v, j) => <td key={j} style={{ border: `1px solid ${LINE}`, padding: "3px 4px" }}>{formatStatValue(otherLabels[j], v)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>

      <h2 style={{ fontSize: 16, marginBottom: 8 }}>Advanced — Four Factors & Ratings</h2>
      <TeamAdvancedStats advanced={advanced} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Team Play — Offense / Defense</h2>
      <OffenseDefenseBreakdown off={teamOff} def={teamDef} />

      <h2 style={{ fontSize: 16, marginTop: 24, marginBottom: 8 }}>Resources</h2>
      {resources.length === 0 ? <p>No resource shared.</p> : (
        <ul style={{ fontSize: 12.5, paddingLeft: 18 }}>
          {resources.map(r => <li key={r.id} style={{ marginBottom: 4 }}>{r.title} — {r.type} — {r.url}</li>)}
        </ul>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Resources — liens vidéo ou documents partagés à toute l'équipe, importés par le coach.
// Les deux plus récents apparaissent sur l'écran d'accueil de chaque joueur, juste sous Role.
// ---------------------------------------------------------------------------

function detectResourceType(url) {
  return /youtube\.com|youtu\.be|vimeo\.com|\.mp4($|\?)/i.test(url) ? "video" : "document";
}

// ---------------------------------------------------------------------------
// Planning — emploi du temps de l'équipe. Le coach ajoute des événements (entraînement,
// meeting, autre) avec heure de début/fin, lieu et couleur ; le joueur consulte en
// lecture seule. L'écran d'accueil affiche uniquement les événements du jour même.
// ---------------------------------------------------------------------------

const DEFAULT_PLANNING_EVENT_TYPES = ["Training", "Meeting", "Other"];
const PLANNING_COLORS = ["#F2A93B", "#2FBF9C", "#E4231C", "#4A90D9", "#B15FE0", "#8B93A1"];

function startOfWeek(d) {
  const date = new Date(d);
  const day = date.getDay(); // 0 = dimanche
  const diff = day === 0 ? -6 : 1 - day; // ramène au lundi de la semaine
  date.setDate(date.getDate() + diff);
  date.setHours(0, 0, 0, 0);
  return date;
}
function addDays(d, n) { const c = new Date(d); c.setDate(c.getDate() + n); return c; }
function toDateKey(d) { return todayLocal(d); }
function formatDayLabel(d) { return d.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" }); }

function PlanningTab({ isCoach, team, roster = [], playerName }) {
  const [events, setEvents] = useState([]);
  const [eventTypes, setEventTypes] = useState(DEFAULT_PLANNING_EVENT_TYPES);
  const [loading, setLoading] = useState(true);
  const [weekStart, setWeekStart] = useState(startOfWeek(new Date()));
  const [editingId, setEditingId] = useState(null); // null = nouvel événement, sinon en cours de modification
  const [date, setDate] = useState(todayLocal());
  const [startTime, setStartTime] = useState("09:00");
  const [endTime, setEndTime] = useState("10:00");
  const [type, setType] = useState(DEFAULT_PLANNING_EVENT_TYPES[0]);
  const [title, setTitle] = useState("");
  const [location, setLocation] = useState("");
  const [color, setColor] = useState(PLANNING_COLORS[0]);
  const [playerIds, setPlayerIds] = useState([]); // vide = concerne toute l'équipe
  const [viewAsPlayerId, setViewAsPlayerId] = useState(""); // coach : voir le planning d'un joueur précis
  const [newTypeName, setNewTypeName] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [requestedIds, setRequestedIds] = useState(new Set());
  const formRef = useRef();

  useEffect(() => { load(); }, []);
  async function load() {
    setLoading(true);
    setEvents((await storeGet("planning_events")) || []);
    setEventTypes((await storeGet("planning_event_types")) || DEFAULT_PLANNING_EVENT_TYPES);
    setLoading(false);
  }

  function togglePlayer(id) {
    setPlayerIds(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  }

  async function addEventType() {
    const name = newTypeName.trim();
    if (!name || eventTypes.includes(name)) return;
    const next = [...eventTypes, name];
    await storeSet("planning_event_types", next);
    setEventTypes(next);
    setNewTypeName("");
  }

  function resetForm() {
    setEditingId(null);
    setDate(todayLocal()); setStartTime("09:00"); setEndTime("10:00");
    setType(eventTypes[0]); setTitle(""); setLocation(""); setColor(PLANNING_COLORS[0]); setPlayerIds([]);
  }

  // Pré-remplit le formulaire avec les données d'un événement — pour le modifier (editingId
  // défini) ou pour le dupliquer sur un autre jour (editingId reste vide, donc "Add event"
  // créera une NOUVELLE entrée au lieu d'écraser l'originale).
  function loadIntoForm(ev, forEditing) {
    setEditingId(forEditing ? ev.id : null);
    setDate(ev.date); setStartTime(ev.startTime); setEndTime(ev.endTime);
    setType(ev.type); setTitle(ev.title); setLocation(ev.location || "");
    setColor(ev.color); setPlayerIds(ev.playerIds || []);
    setError("");
    if (formRef.current) formRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function submitEvent() {
    setError("");
    if (!title.trim()) { setError("Add a title."); return; }
    if (endTime <= startTime) { setError("End time must be after start time."); return; }
    setBusy(true);
    if (editingId) {
      const next = events.map(ev => ev.id === editingId
        ? { ...ev, date, startTime, endTime, type, title: title.trim(), location: location.trim(), color, playerIds }
        : ev);
      await storeSet("planning_events", next);
      setEvents(next);
    } else {
      const entry = { id: uid(), date, startTime, endTime, type, title: title.trim(), location: location.trim(), color, playerIds };
      const next = [...events, entry];
      await storeSet("planning_events", next);
      setEvents(next);
    }
    resetForm();
    setBusy(false);
  }

  async function handleDelete(ev) {
    await requestDeletion(team.id, team.name, "planning_event", `${ev.title} — ${ev.date}`, { id: ev.id });
    setRequestedIds(s => new Set([...s, ev.id]));
    setConfirmDeleteId(null);
  }

  if (loading) return <EmptyState text="Loading…" />;

  // Un joueur ne voit que les événements qui le concernent : ceux sans sélection de
  // joueurs (= toute l'équipe) et ceux où il est explicitement listé. Le coach voit tout
  // par défaut, mais peut choisir de consulter le planning d'un joueur précis.
  const me = !isCoach && playerName ? roster.find(p => p.name === playerName) : null;
  const viewingAs = isCoach && viewAsPlayerId ? roster.find(p => p.id === viewAsPlayerId) : me;
  const visibleEvents = (!isCoach || viewAsPlayerId)
    ? events.filter(ev => !ev.playerIds || ev.playerIds.length === 0 || (viewingAs && ev.playerIds.includes(viewingAs.id)))
    : events;

  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
  const byDay = {};
  for (const d of days) byDay[toDateKey(d)] = [];
  for (const ev of visibleEvents) { if (byDay[ev.date]) byDay[ev.date].push(ev); }
  for (const k in byDay) byDay[k].sort((a, b) => a.startTime.localeCompare(b.startTime));

  return (
    <div>
      {isCoach && (
        <div ref={formRef} style={{ background: PANEL, border: `1px solid ${editingId ? AMBER : LINE}`, borderRadius: 12, padding: 18, marginBottom: 24 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div style={{ fontSize: 14, fontWeight: 700 }}>{editingId ? "Edit event" : "New event"}</div>
            {editingId && <button onClick={resetForm} style={{ fontSize: 11.5, color: "#8B93A1", background: "none", border: "none", cursor: "pointer" }}>Cancel edit</button>}
          </div>
          <div style={{ marginBottom: 16 }}>
            <label style={labelStyle}>Date</label>
            <input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%", maxWidth: 320, height: 46, boxSizing: "border-box" }} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", columnGap: 20, marginBottom: 16, boxSizing: "border-box" }}>
            <div style={{ minWidth: 0 }}>
              <label style={labelStyle}>Start</label>
              <input type="time" value={startTime} onChange={e => setStartTime(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%", height: 46, boxSizing: "border-box" }} />
            </div>
            <div style={{ minWidth: 0 }}>
              <label style={labelStyle}>End</label>
              <input type="time" value={endTime} onChange={e => setEndTime(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%", height: 46, boxSizing: "border-box" }} />
            </div>
          </div>
          <div style={{ marginBottom: 16 }}>
            <label style={labelStyle}>Type</label>
            <select value={type} onChange={e => setType(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%", maxWidth: 260 }}>
              {eventTypes.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
            <div style={{ display: "flex", gap: 8, marginTop: 8, maxWidth: 320 }}>
              <input value={newTypeName} onChange={e => setNewTypeName(e.target.value)} placeholder="New category…" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", fontSize: 12, padding: "8px 10px", flex: 1 }} />
              <button type="button" onClick={addEventType} style={{ ...btnSecondary, padding: "8px 14px", fontSize: 12, flexShrink: 0 }}>Add</button>
            </div>
          </div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 16 }}>
            <div style={{ flex: "1 1 180px", minWidth: 180 }}>
              <label style={labelStyle}>Title</label>
              <input value={title} onChange={e => setTitle(e.target.value)} placeholder="e.g. Full team practice" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%" }} />
            </div>
            <div style={{ flex: "1 1 180px", minWidth: 180 }}>
              <label style={labelStyle}>Location</label>
              <input value={location} onChange={e => setLocation(e.target.value)} placeholder="e.g. Main gym" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", width: "100%" }} />
            </div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>Color</label>
            <div style={{ display: "flex", gap: 8 }}>
              {PLANNING_COLORS.map(c => (
                <button key={c} type="button" onClick={() => setColor(c)} style={{
                  width: 26, height: 26, borderRadius: "50%", background: c, cursor: "pointer",
                  border: color === c ? `2px solid ${PAPER}` : "2px solid transparent", padding: 0,
                }} />
              ))}
            </div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>Players concerned (leave empty for the whole team)</label>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {roster.map(p => (
                <button key={p.id} type="button" onClick={() => togglePlayer(p.id)} style={{
                  fontSize: 12, padding: "5px 10px", borderRadius: 20, cursor: "pointer",
                  background: playerIds.includes(p.id) ? AMBER : PANEL2,
                  color: playerIds.includes(p.id) ? "#1A1300" : "#8B93A1",
                  border: `1px solid ${playerIds.includes(p.id) ? AMBER : LINE}`,
                }}>{p.name}</button>
              ))}
            </div>
          </div>
          {error && <div style={{ color: RED, fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          <button disabled={busy} onClick={submitEvent} style={{ ...btnPrimary, width: "auto", padding: "9px 18px" }}>{busy ? "…" : editingId ? "Save changes" : "Add event"}</button>
        </div>
      )}

      {isCoach && (
        <div style={{ marginBottom: 16 }}>
          <label style={labelStyle}>View schedule for</label>
          <select value={viewAsPlayerId} onChange={e => setViewAsPlayerId(e.target.value)} style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit", maxWidth: 260 }}>
            <option value="">— Whole team (everything) —</option>
            {roster.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
        </div>
      )}

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <button onClick={() => setWeekStart(addDays(weekStart, -7))} style={btnSecondary}><ChevronLeft size={15} /></button>
        <div style={{ fontSize: 13.5, fontWeight: 700 }}>{formatDayLabel(days[0])} – {formatDayLabel(days[6])}</div>
        <button onClick={() => setWeekStart(addDays(weekStart, 7))} style={btnSecondary}><ChevronRight size={15} /></button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        {days.map(d => {
          const key = toDateKey(d);
          const isToday = key === toDateKey(new Date());
          return (
            <div key={key}>
              <div style={{ fontSize: 12, fontWeight: 700, color: isToday ? AMBER : "#8B93A1", textTransform: "uppercase", marginBottom: 6 }}>
                {formatDayLabel(d)}{isToday ? " · Today" : ""}
              </div>
              {byDay[key].length === 0 ? (
                <div style={{ fontSize: 12.5, color: "#5C6470", paddingLeft: 4 }}>—</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {byDay[key].map(ev => (
                    <div key={ev.id} style={{ display: "flex", alignItems: "center", gap: 10, background: PANEL, border: `1px solid ${LINE}`, borderLeft: `4px solid ${ev.color}`, borderRadius: 8, padding: "8px 12px" }}>
                      <div style={{ fontSize: 12, color: "#8B93A1", width: 90, flexShrink: 0 }}>{ev.startTime}–{ev.endTime}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13.5, fontWeight: 600 }}>{ev.title} <span style={{ color: "#5C6470", fontWeight: 400 }}>· {ev.type}</span></div>
                        {ev.location && <div style={{ fontSize: 11.5, color: "#5C6470" }}>{ev.location}</div>}
                        {isCoach && (
                          <div style={{ fontSize: 11, color: "#5C6470" }}>
                            {!ev.playerIds || ev.playerIds.length === 0 ? "Whole team" : ev.playerIds.map(id => roster.find(p => p.id === id)?.first).filter(Boolean).join(", ")}
                          </div>
                        )}
                      </div>
                      {isCoach && (
                        requestedIds.has(ev.id) ? (
                          <span style={{ fontSize: 11, color: AMBER, flexShrink: 0 }}>Pending admin approval</span>
                        ) : confirmDeleteId === ev.id ? (
                          <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                            <button onClick={() => handleDelete(ev)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                            <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                          </div>
                        ) : (
                          <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                            <button onClick={() => loadIntoForm(ev, true)} title="Edit" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><ClipboardList size={14} /></button>
                            <button onClick={() => loadIntoForm(ev, false)} title="Duplicate to another day" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Plus size={14} /></button>
                            <button onClick={() => setConfirmDeleteId(ev.id)} title="Delete" style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex" }}><Trash2 size={14} /></button>
                          </div>
                        )
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function TeamResourcesTab({ isCoach, team }) {
  const [resources, setResources] = useState([]);
  const [loading, setLoading] = useState(true);
  const [title, setTitle] = useState("");
  const [url, setUrl] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [requestedIds, setRequestedIds] = useState(new Set());

  useEffect(() => { load(); }, []);
  async function load() {
    setLoading(true);
    setResources(((await storeGet("team_resources")) || []).sort((a, b) => b.addedAt.localeCompare(a.addedAt)));
    setLoading(false);
  }

  async function addResource() {
    setError("");
    if (!title.trim() || !url.trim()) { setError("Add a title and a link."); return; }
    let parsed;
    try { parsed = new URL(url.trim()); } catch { setError("That doesn't look like a valid link."); return; }
    setBusy(true);
    const entry = { id: uid(), title: title.trim(), url: parsed.href, type: detectResourceType(parsed.href), addedAt: new Date().toISOString() };
    const next = [entry, ...resources];
    await storeSet("team_resources", next);
    setResources(next);
    setTitle(""); setUrl("");
    setBusy(false);
  }

  async function handleDelete(r) {
    await requestDeletion(team.id, team.name, "resource", r.title, { id: r.id });
    setRequestedIds(s => new Set([...s, r.id]));
    setConfirmDeleteId(null);
  }

  if (loading) return <EmptyState text="Loading…" />;

  return (
    <div>
      {isCoach && (
        <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, marginBottom: 20 }}>
          <p style={{ color: "#8B93A1", fontSize: 13, lineHeight: 1.6, margin: "0 0 14px" }}>
            Share a video link (YouTube, Vimeo...) or a document link (PDF, Google Drive...) with the whole team.
            The two most recent show up on every player's Home screen.
          </p>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
            <div style={{ flex: 1, minWidth: 180 }}>
              <label style={labelStyle}>Title</label>
              <input value={title} onChange={e => setTitle(e.target.value)} placeholder="e.g. Zone offense breakdown" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
            </div>
            <div style={{ flex: 2, minWidth: 220 }}>
              <label style={labelStyle}>Link</label>
              <input value={url} onChange={e => setUrl(e.target.value)} placeholder="https://…" style={{ ...inputStyle, letterSpacing: "normal", fontFamily: "inherit" }} />
            </div>
          </div>
          {error && <div style={{ color: RED, fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          <button disabled={busy} onClick={addResource} style={{ ...btnPrimary, width: "auto", padding: "9px 18px" }}>{busy ? "…" : "Add"}</button>
        </div>
      )}

      {resources.length === 0 ? <EmptyState text="No resource shared yet." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {resources.map(r => (
            <div key={r.id} style={{ ...btnRow, cursor: "default" }}>
              <a href={r.url} target="_blank" rel="noopener noreferrer" style={{ display: "flex", alignItems: "center", gap: 10, color: "inherit", textDecoration: "none", flex: 1, minWidth: 0 }}>
                {r.type === "video" ? <Video size={16} color={AMBER} /> : <LinkIcon size={16} color={TEAL} />}
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.title}</span>
              </a>
              {isCoach && (
                requestedIds.has(r.id) ? (
                  <span style={{ fontSize: 11.5, color: AMBER, flexShrink: 0 }}>Pending admin approval</span>
                ) : confirmDeleteId === r.id ? (
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                    <button onClick={() => handleDelete(r)} style={{ background: RED, border: "none", borderRadius: 6, color: "#fff", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Yes</button>
                    <button onClick={() => setConfirmDeleteId(null)} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 6, color: "#8B93A1", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>Cancel</button>
                  </div>
                ) : (
                  <button onClick={() => setConfirmDeleteId(r.id)} style={{ background: "none", border: "none", color: "#5C6470", cursor: "pointer", display: "flex", flexShrink: 0 }}><Trash2 size={14} /></button>
                )
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}


function TeamTab({ roster, allPlays, matchesIndex, matchFilter, isCoach, team, visibility }) {
  const v = (visibility || DEFAULT_VISIBILITY).team;
  const box = useAllBoxScores(matchFilter);
  const advanced = useTeamAdvancedStats(matchFilter);
  const TEAM_SUBTAB_ORDER = [["classement", "standings"], ["collectif", "teamPlay"], ["avance", "advanced"], ["resources", "resources"]];
  const defaultTeamSubtab = isCoach ? "classement" : ((TEAM_SUBTAB_ORDER.find(([, key]) => v[key]) || ["classement"])[0]);
  const [subtab, setSubtab] = useState(defaultTeamSubtab);
  const [exportReport, setExportReport] = useState(null);

  // Colonnes de stats communes à afficher (hors points, déjà en tête) : les 3 plus fréquentes
  // parmi tous les joueurs ayant un box score.
  const otherLabels = useMemo(() => {
    const freq = {};
    Object.values(box.byPlayer).forEach(b => {
      b.statLabels.forEach(l => { if (l !== b.ptsLabel) freq[l] = (freq[l] || 0) + 1; });
    });
    return Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([l]) => l);
  }, [box.byPlayer]);

  const rows = roster.map(p => {
    const b = box.byPlayer[p.name];
    const codedGames = new Set(allPlays.filter(pl => pl.player === p.name).map(pl => pl.matchId)).size;
    return {
      name: p.name,
      games: b ? b.games : 0,
      codedGames,
      ppg: b && b.ptsLabel && b.averages[b.ptsLabel] !== null ? b.averages[b.ptsLabel] : null,
      others: otherLabels.map(l => b && b.averages[l] !== undefined ? b.averages[l] : null),
    };
  }).filter(r => r.games > 0).sort((a, b) => (b.ppg ?? -1) - (a.ppg ?? -1));

  // Bloc collectif : toutes les actions de tous les joueurs confondues, attaque et défense.
  const teamOff = useMemo(() => allPlays.filter(isOffense), [allPlays]);
  const teamDef = useMemo(() => allPlays.filter(isDefense), [allPlays]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <SectionTitle eyebrow="Overview" title="Team" />
        {rows.length > 0 && (
          <button onClick={async () => {
            const filename = `team_${team?.name || "export"}_${todayLocal()}.html`;
            const pdfOk = await tryExportPdf("team-print-content", filename, "light");
            if (pdfOk) return;
            const r = buildReportHtml("team-print-content", filename, "light");
            if (!r) { alert("Content not found — try again after the page has fully loaded."); return; }
            tryDownload(r.full, r.filename);
            setExportReport(r);
          }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: PANEL2, border: `1px solid ${LINE}`, borderRadius: 8, color: PAPER, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
            <Download size={14} /> Export PDF
          </button>
        )}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <StatPill label="Coded matches" value={new Set(allPlays.map(p => p.matchId)).size} />
        <StatPill label="Box scores imported" value={new Set(Object.values(box.byPlayer).flatMap(b => b.entries.map(e => e.date + e.opponent))).size} tone="teal" />
        <StatPill label="Players with box score" value={rows.length} />
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 20, borderBottom: `1px solid ${LINE}`, paddingBottom: 10, flexWrap: "wrap" }}>
        {[["classement", "Standings", "standings"], ["collectif", "Team Play", "teamPlay"], ["avance", "Advanced", "advanced"], ["resources", "Resources", "resources"]]
          .filter(([, , key]) => isCoach || v[key])
          .concat(isCoach ? [["entrainement", "Team Training"]] : [])
          .map(([id, label]) => (
          <button key={id} onClick={() => setSubtab(id)} style={{
            padding: "7px 12px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "inherit",
            background: subtab === id ? PANEL2 : "transparent", color: subtab === id ? AMBER : "#8B93A1"
          }}>{label}</button>
        ))}
      </div>

      {!isCoach && !v[{ classement: "standings", collectif: "teamPlay", avance: "advanced", resources: "resources" }[subtab]] && (
        <div style={{ padding: 30, textAlign: "center", color: "#5C6470", border: `1px dashed ${LINE}`, borderRadius: 12, fontSize: 13.5, marginBottom: 26 }}>
          You don't have the permission to see this.
        </div>
      )}

      {subtab === "classement" && (isCoach || v.standings) && (
        rows.length === 0 ? (
          <EmptyState text="Import box scores ('Full Stats' tab) to display the team standings — this is the source of truth for totals." />
        ) : (
          <div style={{ background: PANEL, border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
            <div style={{ display: "grid", gridTemplateColumns: `0.3fr 1.6fr 0.7fr 0.8fr ${otherLabels.map(() => "0.8fr").join(" ")}`, padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5C6470", borderBottom: `1px solid ${LINE}` }}>
              <div>#</div><div>Joueur</div><div>Matchs</div><div>Pts/match</div>
              {otherLabels.map(l => <div key={l}>{l}</div>)}
            </div>
            {rows.map((r, i) => (
              <div key={r.name} style={{ display: "grid", gridTemplateColumns: `0.3fr 1.6fr 0.7fr 0.8fr ${otherLabels.map(() => "0.8fr").join(" ")}`, padding: "10px 16px", alignItems: "center", borderBottom: `1px solid ${LINE}`, fontSize: 13.5 }}>
                <div style={{ fontFamily: "ui-monospace, monospace", color: "#5C6470" }}>{i + 1}</div>
                <div style={{ fontWeight: 600 }}>{r.name}</div>
                <div style={{ color: "#8B93A1" }}>{r.games}</div>
                <div style={{ fontFamily: "ui-monospace, monospace", color: AMBER, fontWeight: 700 }}>{r.ppg !== null ? r.ppg.toFixed(1) : "–"}</div>
                {r.others.map((v, j) => <div key={j} style={{ fontFamily: "ui-monospace, monospace", color: "#8B93A1" }}>{formatStatValue(otherLabels[j], v)}</div>)}
              </div>
            ))}
          </div>
        )
      )}

      {subtab === "collectif" && (isCoach || v.teamPlay) && <OffenseDefenseBreakdown off={teamOff} def={teamDef} detailTables={false} />}

      {subtab === "avance" && (isCoach || v.advanced) && <TeamAdvancedStats advanced={advanced} />}
      {subtab === "resources" && (isCoach || v.resources) && <TeamResourcesTab isCoach={isCoach} team={team} />}
      {subtab === "entrainement" && isCoach && <CollectiveTraining roster={roster} />}

      {rows.length > 0 && (
        <div className="print-only" id="team-print-content">
          <TeamPrintReport team={team} rows={rows} otherLabels={otherLabels} advanced={advanced} teamOff={teamOff} teamDef={teamDef} roster={roster} />
        </div>
      )}
      <ExportModal report={exportReport} onClose={() => setExportReport(null)} />
    </div>
  );
}
