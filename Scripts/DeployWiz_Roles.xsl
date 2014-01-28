<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

  <xsl:template match="Roles">

    <div>
      <xsl:for-each select="Role | RoleService | Feature">
        <xsl:apply-templates select="." />
      </xsl:for-each>
    </div>

  </xsl:template>

  <xsl:template match="Role | RoleService | Feature">

    <div class="TreeDir">
      <input type="checkbox" class="TreeDirLabel">
        <xsl:choose>
          <xsl:when test="@Id">
            <xsl:attribute name="id">
              <xsl:value-of select="name()"/>.<xsl:value-of select="@Id"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="disabled"></xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:value-of select="@DisplayName"/>
      </input>

      <!-- Parse Each Child -->
      <xsl:for-each select="Role | RoleService | Feature">
        <xsl:apply-templates select="." />
      </xsl:for-each>

    </div>
  </xsl:template>

</xsl:stylesheet>
