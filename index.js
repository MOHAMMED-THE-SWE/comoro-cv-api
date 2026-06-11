const express = require('express');
const cors = require('cors');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  ImageRun, Header, Footer, PageNumber, LevelFormat, UnderlineType,
  TabStopType, TabStopPosition
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ── COMORO BRAND COLOURS ──
const BLUE = "1B4F8C";
const DARK = "1F2937";
const GRAY = "6B7280";
const WHITE = "FFFFFF";
const LIGHT_BLUE = "EBF2FA";

// ── LOGO BASE64 ──
const LOGO_BASE64 = "/9j/4AAQSkZJRgABAQAAlgCWAAD/4QCMRXhpZgAATU0AKgAAAAgABQESAAMAAAABAAEAAAEaAAUAAAABAAAASgEbAAUAAAABAAAAUgEoAAMAAAABAAIAAIdpAAQAAAABAAAAWgAAAAAAAACWAAAAAQAAAJYAAAABAAOgAQADAAAAAQABAACgAgAEAAAAAQAAAWagAwAEAAAAAQAAAIEAAAAA/+0AOFBob3Rvc2hvcCAzLjAAOEJJTQQEAAAAAAAAOEJJTQQlAAAAAAAQ1B2M2Y8AsgTpgAmY7PhCfv/AABEIAIEBZgMBIgACEQEDEQH/xAAfAAABBQEBAQEBAQAAAAAAAAAAAQIDBAUGBwgJCgv/xAC1EAACAQMDAgQDBQUEBAAAAX0BAgMABBEFEiExQQYTUWEHInEUMoGRoQgjQrHBFVLR8CQzYnKCCQoWFxgZGiUmJygpKjQ1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4eLj5OXm5+jp6vHy8/T19vf4+fr/xAAfAQADAQEBAQEBAQEBAAAAAAAAAQIDBAUGBwgJCgv/xAC1EQACAQIEBAMEBwUEBAABAncAAQIDEQQFITEGEkFRB2FxEyIygQgUQpGhscEJIzNS8BVictEKFiQ04SXxFxgZGiYnKCkqNTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqCg4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2dri4+Tl5ufo6ery8/T19vf4+fr/2wBDAAICAgICAgMCAgMFAwMDBQYFBQUFBggGBgYGBggKCAgICAgICgoKCgoKCgoMDAwMDAwODg4ODg8PDw8PDw8PDw//2wBDAQICAgQEBAcEBAcQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/3QAEABf/2gAMAwEAAhEDEQA/AP38ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiobi4gtIJLm6kWKGJSzu5CqqgZJJPAAFeD3vxE8U+OryXRfhVagW0R2zarcLtiU+kYI6/UE/7I61wY7MqdCylrJ7Jat+i/XbzM6lVR3Pcb/UtO0q3N3qd1FaQDgvM6xr+bECvP7z4x/Dezco+tRyMOP3au4/NVIP51zum/BLR7iddT8cahc+JNRPJaZ2WFc84VAc4+px7CvQIPBngrSYM2+j2MKoPvNCnH1ZgT+tcPtcwqLmjGMF5tyf4WX4she1lskvxNbQdf0rxNpkesaNN9otJSwV8FeVOCMHB61s15PZ3nxBtYvK0610u4gi/5ZwNtC55wAGAFXV+IM2musPirSLjTC3AlX95ET/vDH5DNeFR4+wkIr67GdLznTlGN/8AFrFfNnu/2LX2g1J+TV/u3PS6KpWGpWOqWy3mnTLcQt0ZTn8D3B9jzV2vtaNaFSCqU5Jxeqa1T9GeXODi+WSswooorUkKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP//Q/fyig1+bP7Y37a7fC25uPhh8KpY5/FYAW+vmUSRaduGQiA5V58EE5BVAectwOzAYCpiqipUlr+RzYvFwow56j0PvLxp8RvAfw508ap498QWOg2z5CNeTpEZCOSsasdztjsoJr5wvv28/2X7OXyk8VvdYzlobG6K8e7RLn8K/nx8R+JfEHi7V5tf8U6lcatqNwcyXF1I0sje25iSB6DoK1tE+HXxB8S2xvfDnhfVdVtx1ktLGe4QY/wBqNGFfd0eDKEI3r1Hfysl+Nz5OpxNVlK1KC/M/ox8GftZ/s7ePLtNO0HxvYpeSkKkN4Xsndm4Cp9pWMOx9FJNfRKsGG5TkHoRX8kWp6TqekXL2GsWc1jcJkNDcRtE49irgH9K/SL/gnx8VfjhdfESD4c6VI+ueDEhaW/jvHZk06IDCSQSnJRi+FWLlXyeFwXXgzXhKNKk61Gei11/zOzL+IZVKip1Y6vt/kft5TWYKCzHAHUntXHfEHx94Z+GPg/U/HPjC6FnpWlxGSRurMTwkaD+J3YhVHckV/P1+0L+1/wDEr466ldWMV1LoHhPJWHS7aQr5if3rl1wZXbuPuDoB1J8HJ8jq4yXu6RW7/rc9bMc1p4Ze9q30P2v8ZftY/s7+A7yTTtf8cWLXkRKvDZ7710YdVf7MsgQ+zEVwdn+3p+y/dzeVJ4qkth/flsbrb/47Gx/Sv55NO0zUNUuUsdJtJbyd+FigjaRz9FUEn8q6nV/hr8RvD1p9v1/wpq2mW3/PW6sLiCP/AL6kQDv619jHg7CRtGc3f1S/Q+afEuIesYK3zP6gPBPxK+H/AMSLA6l4D8Q2Ou26gbzaTpK0ZPQSIDuQ+zAH2rt6/ks0DxFr3hTV4Ne8M6hcaXqNscx3FtI0Uq/RlIOPUdDX7T/scftsyfE+8tfhb8V5I4vFDqVsNQAEceobBny5FHCz4BIIwr44AbhvCzfhSph4urSfNFb91/metl3EEKslCorP8D9LK+br/wDa9/Zu0u+uNM1Dx1Zw3NpI8UqFJ8o6EqynEZHBGK+ka/lJ+Jf/ACUXxT/2FL3/ANHPXLw7k1PGSmqjatbY6M5zOeGjFwSd+5/TH8O/jb8KfixNeW3w68S2utzaeqvPHCWWSNHOFYo6q20kYyBjPHWvU6/l8/Z7+LV58E/i1oXj6F2FnbyeRfwpn97YzELMhA64GHUf3lU9q/p4sb201Oyt9RsJluLW6jSWKRCGV45AGVlI4IIIINZcQZL9TqJRd4vb9S8ozP6zBuSs0W6KKo6nqVjo2nXWr6pOlrZWMTzzzSHakcUSlndj2CqCSfSvBSvoeu2eb/ET44/Cb4T3VpY/EPxPaaJdXyGSGGUs0rRqcF9kYZgucgMQASCAcg1xGlftcfs465qlnouleObO4vtQnitreJUnBkmmYIiDMYGWYgc1+Avx++LN38a/ixrvj6belrdy+VYxPwYrOH5YVI7Er8zD+8x69awfg3/yWDwH/wBjBpP/AKWRV+g0uDaao89ST5rXex8dU4ln7XlhFWv5n9UtNZlQFmIAHJJ7CnV4z8Y9e1CLTrHwXoTbdT8Syi2UjqkOQJD7ZyAT2BJ461+Y5hjI4ejKrJXt07vovm9D66rU5YuTOVvZdQ+NviGbSLCZ7bwbpUgFxKpw17KpztX/AGfT0HzHkgD1PxB4j8GfCjwwtzqDR6dp9uPLhhjHzyPjO2NerMepP1JPU1PbQeHvhh4JO9hb6bo1u0kr8ZcqMsx9WdunqSAK+J/EvjeOC2T42/EWyOq6jqkjW3hLw/gsrndhZXQcsoJBJx8xIIHMePlsRXlg1d2eImrtvaK/PlWyS1kz7DgrhCeYVeeom1dKy3lJ6qEb6LS7lJ6RinJnofiLx/8AEbxPpT+KNW1a3+F/gxv9VcXRBvLhTyCi8MxYcqqbcjoW4NeB/wDCV/AjU7xkg07xd8Tb2P78saOV+qqpWQL6ZNZHjW+sPC/iKx1H43W1z8TvivrKq+neE7NibWwWX7gmEYZRz/CARgEkNw9etaT8NP2yvGtkk2t+LtK+FmlEbo9K0a2RmtlPYsnfHX98RnpjoPSwXAtbFQWJx01GL2dS8m/SC0S+T9T9LxPFGV5V+4wycmt/Zy9nBek7OrU85NxT6Kx59e618FNJ2y+IPhr4y8FhT/x9skqlCemDMxGcY7d+9eo+CvEeu6lBJL8F/HcPje2hTdNoesDbeeWOo2ykFgOBuUqueOa9n+DPjzwbZeF7Hwbq/wASYPGWtRtIGvLv/RpbjzHLKgSQ/NtB2j5mJxTfiZ+zV4H8bka94ZX/AIRTxTbnzbbUtPHknzB084UDA9yMsOxxkHlzTgWeElKMI2fknTbXl0d+0o27nlZT4n5NmiVOpJqO3M5+3h/28pLnS84T5l2exieDfFun+JLy5l8HRS+HfE9kCb3RLrID7PvbNwXdg8YwCPRep+gvC/ii08S2bSxqYbqH5Z4W+8jfjjj/wDV1r4VsLrxP4s8Qy/Dvx0E0T4veGUE2l6pEfLTV4I1JALDALFeQccgHIGHWva/CnjaXXtMi8fW0P2XWtJlFlrtnjbl+m/aOgbHXs3+6c/nuHxNTJarxFH+Dq5wSsrL4pxj9icN5xXuyj7y1QuLOFFa8fK2vN8XwtS+3TntGT1jL3Zan1PQar2tzDeW8V3btuimUOp9QwyKsV+2U6kZRUou6Z+PNNOzPnvxB+1Z+z34V1y/8NeIfGtpZanpkz29zA6TFo5YzhlO2MjIPoTWP/w2Z+zF/wBD/Zf9+5//AI1X4R/tO/8AJxHxG/7Dl7/6MNeZeGPBPjTxvcT2ngrw/qPiCe2QPLHp1pNdvGhOAzrCrlQTwCRjNfpFHg/DOlGpKbV0n0/yPi6nEddVHCMVo/M/o5079rj9mvVJ/s9t8QtKjbIGbiVrZOf9uZUXHHJzx3r3TRPEOg+JtPj1bw1qVrq1jLyk9pMk8TfR4yyn86/lj8R/DH4leD7YXvi7wlrGh2zcCW/0+4tUP/ApUUUzwH8RvHPww1pPEHgHWrnRL1SCWgfCSY7SRnKSL7OpHtUVuC6Uo3oVPv1X4FU+JpxlarD+vmf1b0V8xfsp/tC237Q3w3/t66ijtPEGkyC01S3jPyCXaGWaME5EcoyVBzghlydua938Y+L/AA74B8M6j4x8W3qafpOlRGa4mfoqjgAAcszEhVUZLMQACSBXwtfC1KdV0ZL3k7WPq6WIhOCqReh0pYAZPavAPG/7U37P3w8vH03xP42sEvYjte3tma8ljYfwyLbLIUPs+P1Ffi/+0b+2f8RvjZqVzpOhXM3hrwepKRWNvIUmuEB+/dSIcszf3AQijjDHLH48sbC81CdLLTbaS5nfhIoULuT6BVBJr7TL+C7x58TK3kv8z5nF8TWly0I383/kf0O237e37L9xN5T+KZbcf35LG629f9mMn9K+h/AnxV+HHxNtDe+APEljrsaAF1tplaWPPTzIuJIyfRlFfzFan8MfiVolkdR1rwlq+n2gGTNcafcQxgeu90C/rXM6RrGreHtTg1nQrybTtQtW3RXFvI0UqEf3XUgj867KvBmHnH9zUd/k1+FjmhxNWi/3sFb5o/raor8sv2PP247zxnqNj8KvjHOrazcERadq2AguW/hhuAMASnorgAPwG+blv1MFfC5hl1XC1PZ1V/wT6vB4yFeHPTYtFFFcJ1H/0f2X/aA+KA+Dvwh8SeP4gr3mn2xSzRxlWu5iI4Nw4yodgzDIyoOOa/mF1LUb/WNQutX1W4e7vb2V5p5pDueSWQlndj3JJJNfuN/wU01Sey+Bei2EZITUdet4n9xHbzyjP4oK/C0DJA9a/UODMLGOGdXrJ/gv6Z8JxNXbrqn0S/M/Wv8AYP8A2SdA1/Qrf42/E7T49Riumf8AsewuF3w7I22m6lQ8MWYERqwIAG/nK4/XyGGK3iSCBFjjjUKqqAFVRwAAOAAOlcT8L9Ft/Dfw28LaDaoEisNLs4QB0+SFQfzPNd1XwWbZhPE15Tm9Onkj63L8HGjSUYr1POfiT8Jfh58XNDfw/wDEHRINWtmBCO42zQn+9FMuJI2HqpHvWJ8GfgX8PPgP4en8O+ALOSGO7l865uLiTzridwMKXfAGFHCqoCjk4yST7DSGuP61U9n7LmfL26HT7CHP7S2vc/Eb/gpR8XLzXfiHYfCLT7hhpnhuGO6vIlOFe+uV3JvHcxwspXPTefWvhj4P/DHWPjF8R9E+HeiyCCbVptrzFdwghQF5ZSuRnYgJxnk4Fbv7R2r3Gu/H/wCIupXTl3Ov6hCpOc+XbTtBEOfSNFFfYf8AwTD0a1vPi74o1uYBptN0Xy4sjODc3CbmHoQI8fRjX6vF/Ust5obqP4v/AILPz5/7TjbS2b/Bf8A/XD4VfBv4efBnw7B4c8BaVHZRxoFluGAe6uW43PNLgFmYjOOFHRVAAA9PdFkRo3UMrAgg8gg9jTqK/JalWU5OU3ds/RIQjFcsVZH5Rft5fsoeF4vCl38a/hzpsel32lkPq9rbIEhuLdiF+0LGvCvGSC5UAMuWblc1+PVjfXmmXtvqWnzNb3dpIk0MqEq8ckZDKykcgqQCD61/WD4t0Oz8T+FdZ8Naiu611ayubSUHnMc8bRt+jGv5L4GLwxuerKD+Yr9M4Px861GVOo78tvuf/DHw3EeFjSqxnBW5vzR/UP8As7/FFPjJ8HPDXj9iBd31v5d4o/gu7djDOMdgXQsvqpB6Gv5tPiX/AMlG8U/9hS9/9HPX7F/8Ew9RuZ/g94m0uRi0NprjPHk5x51tDuA9Blc/Umvx0+Jf/JRvFP8A2FL3/wBHPUcO4ZUcZiKcdlb9Ss6ruphqM3uziirKFLAgMMjPcZxkfiK/dP8A4J0/Gf8A4Tb4ZXHww1i58zVvBm0W4c/O+nSk+VjPJELZj/2V2DjivgJPgjJ40/YusfixoUIfVPB+q34vFUfPLp0rR7jnv5D/AD4/uFz1AB8S/Z5+LV18Ifi1oXj2N3Flby+RfomT5ljPhZlwPvYGHUf3lU9RXo5rh487D1KcPii396/zX5nHl9WWErQnLaSX3P8AyP6gcivzi/4KMfGtPBfw4tvhVo02NY8YfNc7TzFp0TfOT3/fOAi+qiT05/QOfxDolr4fk8Vz3sS6RFam9a6DAxC2CeYZdwyCuz5sjtX8yn7QPxbvPjZ8WNd8fTB47O6l8qwhfrFZQ/LCpHZio3MP7zHFfFcK5Z7bEe0kvdhr8+n+Z9Pn+O9nR5IvWX5dTxoI7KzKpIXknHTtz+NejfBv/ksHgP8A7GDSf/SyKvpO7+CK+B/2Lbn4p6zAU1rxjqtiLfcMGLTY2cxge8zjzCe6iPHQ5+bPg3/yWDwH/wBjBpP/AKWRV+jrFRq06jhsrr7kfEvDypzhzdbM/qjNeEaLnxN8b9Y1OX54PDVqlpb+iyS/fb68uv0r3ivnH4d6xa6Le+PvEF8rOBqgVggBPLsBjJHdua/mviDF06UqMq8rQi3KTfRQi3+G/wAj9PdKVSpCnFXbZV+PT3HijXPCPwotHKx69d+fdhTybe3+Y59gNzfVRXyxr/xB0uDVvH37SGpRJPpHw9RdC8J2bDMLXh/dRyhehwT5nThWz1RSPqHxJbahc/HC88QfZpRaaV4Tnlt5yjeWJ2dhw/Tdtc8Z6A18ENaQ6n8K/gJ4TmGbXxV4xub28B+68ltOsC7v+APjFHDWCjjc5/faxcm/WNNKy+9yZ+7YassFw/KdJ+9ywhddHWc5zfq4U4w72bXU+3/2WfgbJ4D8Pt8SPHim/wDiL4vH23Uru4+aaAT/ADi3BP3Sox5mOrfLyqrX1rjPB6Vw3hTx/oli7WPEOh6XHKk/hq6+yXBkUKrPg8oQTkZBHODx0ruq+0zSpWnXlKurSfTsmrq3lbY/CaGOp4mPtacrp3V/R2f3NWPKPEnwQ+Fnimzls9S8O2sZkBxLboIJUJ7q8eD15wcg9wa8b8O6p4p+AnjTT/AXiy/k1bwZr8nk6XfznMlpN0EMjdgcgHPy/wAS4AcD67r5/wD2nNLtdQ+Dms3M4zJp7W9xEe6uJVTIPbhiK9nJMynWqRwWJk5U5u2uvK3opR7NP71oz4nifJqWGoTzPBRUK1JOV1pzJauMrbpra+qdmjmf2qfA1xqvgmP4k+HM2/ibwK41K2uIwBJ5MRDTIT3UKN+Dx8pGOTXIeGtdsJviD4Y8a2qrFofxZ0wx3UI4SO+CYYY9S+FHfLMa9PsPjB4b1LV/Cnwi1+C4utb8Y+HP7RZgim3MZj2yK53bgXw+Plxx15r5V8IWet6B8Cvh6dYtprO80DxPKkAnjMchj89pM4YDjdkdxxX5LxnlksNU9pOPRS8nyyUX8pRcovuf1FwLmH9oZKqE3tLljffkrQnJL/typTjJdm33Pur4bXUq6VdaHcndLpNw8PPXYSSP1yPpXo1ePWGr2nhrW/GmrXSsbWyUXEgQAscAscA4GT9a9O0bVbfXNIstaswywX8Mc8YcYYLIoYZAzzg1z+H2JX9m08NKV5U3OH/bsJygvwSR+WZ1y/WZNfaSl96T/Nn8zX7Tv/JxHxG/7Dl7/wCjDX2X/wAEuf8AkpHjf/sE2/8A6Pr40/ad/wCTiPiN/wBhy9/9GGvYv2Jfj14D+AXjDxLrnj03X2bVbGK3h+yQiZt6S7zuBZcDFf0xmFGVTLeSCu3FfofkuDqRhjuaTsrv9T+gu+sbLU7ObT9Rt47q1uEKSxSqHjdGGCrK2QQR1BFfzEftHeC9F+Hnx08a+DfDieVpenag32aMciKOZFmEY/2Y9+wZ5wOea/Vnxp/wUy+Eun6PM3gXRdS1jVSp8lbqNLW2DHoXcO7kA9QF5HcV+LfijxLrXjXxNqnizxDP9q1TWrmS6uJMY3SzNuO1R0GThVHAGAOlePwjlmJoznOqrJrZ9z0uIsbRqxjGm7tH6P8A/BLy+v4/iR4z02PP2K40mKWX083G4VYs++2STH41s/8ABS/4x3V54h0f4J6ROVs9OiTUtTCn79xLkW8Tf9c48yEdD5inqtfS/wDwT9+BGp/Cv4b3vjPxRbta654yaGUQuMNBYwhjApB5VpC7Ow9CgPK1+UH7Xeqzax+0n4+uZ2LGLUWt1J/uQIsajn0C08GqeJzadSOqivx0QsS50MuhB6OT/Dc8o+GvgHW/ij480P4feHVzf65crAjEErGmC0krY52xxhnbHOFOK/pQ+DfwK+HfwO8NW+geCtMjjnVALm/dQ13dSfxPJIctyeiA7VGAoAFfkF/wTU0i1vvj1f6pOoaTTdGuGiyM7WlkiQsPfaSPoTX7w15vGWPm6yw6fupX9WdvDWDiqbrNatiEZBBGRX5k/tyfsmeGdd8Haj8Xvh5psWma/okZuNQgt0CR31qv+sfYvAmjHzAgZZQQcnbj9N6oarpttrGmXmk3q77e+hkgkX1SRSrD8jXy+XY+eGqqpB/8Fdj3sZhIVqbhJH8k0M01vNHcW8jRSxMHR0JVlZTkMpHIIPII5Ff0ufsp/Fu4+M/wS0HxbqcnmavArWOoN033Vt8jSY7eYu2QgDALEDgV/NLcoI7mWNeiuwH4Gv2f/wCCXGpXM3gDxvpLtmC11WCZB6NPAFb8P3Yr9G4ww0Z4T2nWLX46HxnDdZxxHJ0f6H6j0UUV+Vn3p//S+/f+CkPh2bWv2fI9UhXf/YerWt03GSFdZLcn2/1or8Eq/q2+JHgXSfib4D13wDrnFnrlpJbM4GWjZhlJVH96NwrrnjIGa/l58feBvEPw18Y6r4G8VQfZ9T0idoZBghXA+7ImeqOuGU9wRX6XwXjIyoyoPdO/yf8AwT4jibDtVFV6P8z+lr9n3xlY+Pvgr4N8UafIJFuNMt45MfwzQIIZVPoVkRgRXsVfzzfsl/tdal+zzd3Ph3xBazav4O1KXzpbeEr59rMQA00AYhW3AAMhZQcAgg5z+vGjftn/ALMmt6cmoxePLK0DAExXay20ynHIMcqKSR04yPQmvls4yCvRrS5INxb0aVz3stzelVprmkk1vc+oaSvzV+N3/BRv4feGbB9M+C6jxVrLjH2ueKWGwg+ocRyTN7Lhf9vtXo/7I/7YWm/Hu0fwr4wFtpXji1DP5EO5IL2EDJkt1dmbcg+/HuJA+YEjOOSpkWKhR9vKFl+PrbsdMM1oSqeyjK7/AA+8/IP9rfwpc+Dv2kPH2m3CkLe6nLqUbHoyahi6yPUBpCv1BHaveP8Agm940tPDXx4uvD19KI4/FGmTW0WTgG5gdZ0GT6osgHqSAK+pP+CjPwB1DxToln8avCtsbi80CE2+qxIMu1lksk6gcnyWJD/7LbuApz+Nug67q/hjW7HxH4funsdS0yaO4tp4/vRyxNuVhnIOCOQcgjggg1+i4Gccdl/s762s/Jr+rnxeKjLCYzntpe/yZ/WtRX5yfBf/AKKL/C7xTpUFh8W2bwnrsahZJlikm0+eum5GjDvET1KyDaOznt71rX7Z/wCzJodi1/N49srsKCRHZrLdSsewCRIxBPvgepHNfm1bJsVCfI6bv5K59tSzLDzjzKa+89T+M/jqz+Gnwq8VeOLyQRjStPnkiycb7hlKwRj3eUqo+tfyvRoI41jH8IA/Kvuf9rX9sbUPj+IfCHhWzl0jwfZzedtmI+03sq8I8wUlUVckqgJ5OScgAfLfwu+Gvib4ueOdK8A+E4DNfanKFLkHy4IR/rJpCBwka/MT1PQZJAP6Jw3lzweHlOvo3q/JL+mfG53jVia0Y0tItdOuz+Z+0P/AAT a8LXWi/AWy8Q3cZQeItXuZ4CejQW6JbgY/wCus cn8K/FX4l/8AJRvFP/YUvf8A0c9f1D+A/BukfDzwZovgfQU2WGiWsVrFxgsI1wWP8AtMcsfcmv5ePiX/yUbxT/ANhS9/8ARz1w8MYr2+KxFXvb9bHVn1D2WHo0+x+2v/BPTTrHV/2YZNK1KFbm0vNS1GGaJxlXjkVFZSPQgkV+QX7RHwe1D4HfFjWfAtzmSxjf7Rp85/5bWUxLRE8D5l5R+25TjjBr9iv+Ccf/ACbhF/2Fr7/2Ss7/AKKD/wCCg/wO/wCFjfC4fEPRIN2u+Cw8z7R80+nNzOncmPiVfQBwBla4sFs z1ZQfLKKt8+n+R2YnA+1wEJR3iv x6n54Tf tur9oX8ArRSvaxK ggMwBz29K/Wr9i34H3PwM+D0Hh3XYlj8Q6xcPqGoKvRI5Fwg9Cbgb1Hbr3r8i/hN8Cvip8cbiaz+H2gT6lDAf390zLBaREZGJJnKRggkfKSW55AzX9EnwO8C6p8N/hL4Y8C6zdC6vtGs1inlA2qXbLsqg9AhJUe4r1uLcbGNBYdPWX5L/gnmcM4Vzre2a0j+L/yPWaKKK/Mj7k//0/38ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAExXMeKPBvhzxhZiz1+zS4CZ2SfdkjJ/uOOR9Oh7g11FFZVqMKkXCork UIVZJ+meLjwpo3jHTPh7e6h5niLWIpJ7a2Ct80UO7c2SApOVPHJ460tfJn7RmoKupeEPh3YoJ9a8XX4bZ1MCW5K72/Mgf1AJ619bV4OZ4N4KrGjKSlJpNvtfZH2GUZksZSdaMXGKbSv3W7MKKKK84908Qv/ANovwhpXxfX4M6rpGoWV5Jc29rb6hLsWyuZ7qPzY4o33hywXrt/Hoa9vrwT4n/CeXx78VPhxrxvILCx8Mjy5beSHzvNMk+1gFYbQCpH3yRjjBr3mVXjikljQyFVJCAhSxAzgFiFGT3JAHcivVx/1Pk/dRkr3etuvkeRl31v2l6zjK1tEcr4s8c+FPBVhJqPiXU4bGGNGfYzDzXCjJEcfLOcDtXhsvxW8W+LFi/4VX4OudW07fmTVNVIsrHaMfMFkH7w5yMA/8AAT3q7+z94Ia41/xd8YNbgkGoeLbyWXT3cg+VYKQkI4OQXBZ+v8S9qzfir8bfCvwpvdC8Oa/GX1jxPeJaWFumQzgnBflSAuRk5BHuRyPf4b4aWEoyxGJlL2cm+ROy5l/eaT+dvkfCZ3xnUzLGYfAYWkvbJU3e7dpNe9y6NX5brl0TVrq7PR9L8VeLvh5ZvdePbSbxb4bsFMr6npqkXKQqPme5t+BkD7xQZ6nBr1DRPEuhSm3n0XVbXUvNtI7xFtJll/0eVyscoK/wsBgcc9K5yw1Ky1SziubN9yyrkkYxkgEgZ9RXBeNbLxr4a1Pwr4Y+EFx4f8P2d9qEqX39pQPdSJBnf5cJuMmRwGAjAO7dzXj4bN8LDMqmGgpwrU4pyundSWibV1dWukfUYnIcbVymliq0oVKNWSjBxdnFrVp20akrtrv2PSaKKK+kPjQooooAKKKKACiiigAooooAKKKKACiiigAooooA//V/fyiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//Z";

function sp() { return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] }); }
function line() { return new Paragraph({ spacing: { before: 120, after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE } }, children: [new TextRun("")] }); }

function buildDoc(data) {
  const { full_name, contact, sections } = data;

  const logoBuffer = Buffer.from(LOGO_BASE64, 'base64');

  const children = [];

  // ── LOGO ──
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new ImageRun({ data: logoBuffer, transformation: { width: 160, height: 60 }, type: "jpg" })]
  }));

  // ── CANDIDATE NAME ──
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: full_name || "", bold: true, size: 36, color: DARK, font: "Calibri" })]
  }));

  // ── CONTACT LINE ──
  if (contact) {
    const parts = [contact.location, contact.phone, contact.email, contact.linkedin].filter(Boolean);
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: parts.join("  •  "), size: 20, color: GRAY, font: "Calibri" })]
    }));
  }

  children.push(line());

  // ── SECTIONS ──
  for (const section of (sections || [])) {
    if (!section || !section.title) continue;

    // Section heading
    children.push(new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: section.title.toUpperCase(), bold: true, size: 22, color: BLUE, font: "Calibri", underline: { type: UnderlineType.NONE } })]
    }));

    if (section.type === "text" && section.content) {
      children.push(new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: section.content, size: 20, color: DARK, font: "Calibri" })]
      }));
    }

    else if (section.type === "bullets" && Array.isArray(section.items)) {
      for (const item of section.items) {
        const text = typeof item === "string" ? item : (item.heading || "");
        children.push(new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text, size: 20, color: DARK, font: "Calibri" })]
        }));
      }
    }

    else if ((section.type === "experience" || section.type === "education") && Array.isArray(section.items)) {
      for (const item of section.items) {
        if (!item) continue;

        // Company/Institution name in blue
        if (item.subheading) {
          children.push(new Paragraph({
            spacing: { before: 160, after: 40 },
            children: [new TextRun({ text: item.subheading, bold: true, size: 22, color: BLUE, font: "Calibri" })]
          }));
        }

        // Role + dates on same line
        const roleText = [item.heading, item.dates].filter(Boolean).join("  |  ");
        if (roleText) {
          children.push(new Paragraph({
            spacing: { before: 40, after: 40 },
            children: [new TextRun({ text: roleText, bold: true, size: 20, color: DARK, font: "Calibri" })]
          }));
        }

        // Location
        if (item.location) {
          children.push(new Paragraph({
            spacing: { before: 20, after: 40 },
            children: [new TextRun({ text: item.location, size: 19, color: GRAY, font: "Calibri" })]
          }));
        }

        // Description
        if (item.description) {
          children.push(new Paragraph({
            spacing: { before: 40, after: 40 },
            children: [new TextRun({ text: item.description, size: 20, color: DARK, font: "Calibri" })]
          }));
        }

        // Bullets
        for (const bullet of (item.bullets || [])) {
          children.push(new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            spacing: { before: 30, after: 30 },
            children: [new TextRun({ text: bullet, size: 20, color: DARK, font: "Calibri" })]
          }));
        }

        // Key achievements
        if (item.key_achievements && item.key_achievements.length > 0) {
          children.push(new Paragraph({
            spacing: { before: 60, after: 20 },
            children: [new TextRun({ text: "Key achievements:", bold: true, size: 20, color: DARK, font: "Calibri" })]
          }));
          for (const ach of item.key_achievements) {
            children.push(new Paragraph({
              numbering: { reference: "bullets", level: 0 },
              spacing: { before: 30, after: 30 },
              children: [new TextRun({ text: ach, size: 20, color: DARK, font: "Calibri" })]
            }));
          }
        }
      }
    }

    children.push(sp());
  }

  // ── FOOTER ──
  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE } },
        spacing: { before: 120 },
        children: [
          new TextRun({ text: "+44 (0) 1489 660390    sales@comoro.co.uk    www.comoro.co.uk", size: 18, color: DARK, font: "Calibri" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Comoro Ltd, 12-14 Carlton Place, Southampton, SO15 2EA    Company number: 06802075    VAT number: 303 1352 61", size: 16, color: GRAY, font: "Calibri" })
        ]
      })
    ]
  });

  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }]
    },
    styles: {
      default: { document: { run: { font: "Calibri", size: 20, color: DARK } } }
    },
    sections: [{
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } }
      },
      footers: { default: footer },
      children
    }]
  });

  return doc;
}

// ── ROUTE ──
app.post('/build-cv', async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.full_name) {
      return res.status(400).json({ error: 'Invalid JSON — full_name required' });
    }

    const doc = buildDoc(data);
    const buffer = await Packer.toBuffer(doc);

    const filename = `${(data.full_name || 'CV').replace(/\s+/g, '_')}_Comoro.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('Comoro CV Builder API is running'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
