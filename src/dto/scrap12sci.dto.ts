import { ApiProperty } from "@nestjs/swagger";

export class Scrap12SciDto {
  @ApiProperty({ default: "https://gseb.org/" })
  url: string;
}